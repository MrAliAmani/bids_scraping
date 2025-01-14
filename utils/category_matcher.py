from typing import Dict, List, Tuple, Optional, Set
import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import openai
import numpy as np
from fuzzywuzzy import fuzz
import logging
import json
import os
from rich.console import Console
from collections import defaultdict
from sklearn.feature_extraction.text import TfidfVectorizer
import re
import requests

logger = logging.getLogger(__name__)
console = Console()


class CategoryMatcher:
    def __init__(self, api_categories: List[Dict]):
        """Initialize with list of API categories"""
        # Convert API categories to expected format
        self.api_categories = [
            {"category_id": cat["id"], "category_name": cat["name"]}
            for cat in api_categories
        ]

        self.model = SentenceTransformer("paraphrase-MiniLM-L6-v2")
        self.category_embeddings = None
        self._prepare_embeddings()

        # Initialize GPT-4o client
        self.gpt_client = openai.OpenAI(
            base_url="https://models.inference.ai.azure.com",
            api_key=os.environ["GITHUB_TOKEN"],
        )

        # Initialize Groq client
        self.groq_client = openai.OpenAI(
            base_url="https://api.groq.com/openai/v1",
            api_key=os.environ.get("GROQ_API_KEY"),
        )

        # Initialize OpenRouter client
        self.openrouter_client = openai.OpenAI(
            base_url="https://openrouter.ai/api/v1",
            api_key=os.environ.get("OPENROUTER_API_KEY"),
        )

        # Add Ollama client
        self.ollama_client = openai.OpenAI(
            base_url="http://localhost:11434/v1",
            api_key="ollama",  # Ollama doesn't need an API key
        )

        # Configure confidence thresholds
        self.similarity_threshold = 0.75  # Minimum similarity score to trust
        self.method_weights = {
            "similarity": 0.6,  # Higher weight for similarity method
            "llm": 0.4,  # Lower weight for LLM due to lower accuracy
        }

        # Load override rules
        self.override_rules = self._load_override_rules()

        # Add tracking for processed bids
        self.processed_bids = set()  # Track unique bid identifiers

    def _prepare_embeddings(self):
        """Prepare embeddings for all API categories"""
        category_texts = [cat["category_name"] for cat in self.api_categories]
        self.category_embeddings = self.model.encode(category_texts)

    def match_by_similarity(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using enhanced sentence embeddings and cosine similarity"""
        try:
            # Clean and normalize inputs
            title = " ".join(title.lower().split()).strip()
            description = " ".join(description.lower().split()).strip()
            category = " ".join(category.lower().split()).strip()

            # Extract category codes and descriptions from input category
            category_parts = []
            for part in category.split(";"):
                part = part.strip()
                if "-" in part:
                    code, desc = part.split("-", 1)
                    # Clean up code and description
                    code = code.strip().rstrip(
                        "*"
                    )  # Remove trailing asterisks from codes
                    desc = desc.strip().rstrip(";,.")  # Remove trailing punctuation
                    category_parts.extend([code, desc])
                else:
                    category_parts.append(part.strip().rstrip(";,."))

            # Combine text with strategic weighting
            weighted_parts = []

            # Title is most important (4x)
            weighted_parts.extend([title] * 4)

            # Category codes and descriptions (3x)
            weighted_parts.extend(category_parts * 3)

            # Description if available (1x)
            if description:
                weighted_parts.append(description)

            # Create combined text
            combined_text = " ".join(weighted_parts)

            # Get embedding for combined text
            text_embedding = self.model.encode([combined_text])[0]

            # Calculate similarities
            similarities = cosine_similarity(
                [text_embedding], self.category_embeddings
            )[0]

            # Get top 5 matches
            top_indices = np.argsort(similarities)[-5:][::-1]
            top_scores = similarities[top_indices]

            # Check for exact matches in category codes and keywords
            for idx, cat in enumerate(self.api_categories):
                cat_name = cat["category_name"].lower()
                boost_score = 0

                # Check category codes
                for part in category_parts:
                    if part.strip():
                        # Exact code match
                        if part.strip() in cat_name or cat_name in part.strip():
                            boost_score = max(boost_score, 0.3)
                        # Partial code match
                        elif any(word in cat_name for word in part.strip().split()):
                            boost_score = max(boost_score, 0.2)

                # Check title keywords in category name
                title_words = set(title.split())
                cat_words = set(cat_name.split())
                word_overlap = len(title_words & cat_words)
                if word_overlap > 0:
                    boost_score = max(boost_score, 0.1 * word_overlap)

                # Apply boost
                if boost_score > 0:
                    if idx in top_indices:
                        similarities[idx] *= 1 + boost_score
                    else:
                        similarities[idx] = max(top_scores) * (0.8 + boost_score)

            # Recalculate best match after adjustments
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            # Lower confidence threshold but add minimum word match requirement
            min_threshold = 0.3  # Lower base threshold

            # Check word overlap between title and category
            cat_name = self.api_categories[best_idx]["category_name"].lower()
            title_words = set(
                w for w in title.split() if len(w) > 3
            )  # Only consider words longer than 3 chars
            cat_words = set(w for w in cat_name.split() if len(w) > 3)
            word_overlap = len(title_words & cat_words)

            # Adjust threshold based on word overlap
            if word_overlap > 0:
                min_threshold = max(0.25, min_threshold - (0.05 * word_overlap))

            if best_score < min_threshold:
                return None, 0.0

            return self.api_categories[best_idx], float(best_score)

        except Exception as e:
            logger.error(f"Error in similarity matching: {str(e)}")
            return None, 0.0

    def match_by_llm(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using LLM with fallback options"""
        # Format categories and create prompts
        categories_text = "\n".join(
            [
                f'{{"category_id": {cat["category_id"]}, "category_name": "{cat["category_name"]}"}}'
                for cat in self.api_categories
            ]
        )

        system_prompt = """You are an expert procurement data analyst specializing in government bid categorization. 
Your task is to match procurement requests to the most appropriate category from an authorized list.

CONTEXT:
You are part of a system that processes government bids and must categorize them accurately according to standardized categories.

CORE OBJECTIVE:
Analyze procurement details and select the MOST APPROPRIATE category from the provided list, considering:
1. Solicitation Title (Primary factor)
2. Current Category (Strong hint if available)
3. Description (Supporting context)

MATCHING RULES:
1. EXACT MATCHES:
   - If the title/description exactly matches a category name, prioritize that match
   - Consider industry-standard abbreviations and variations

2. SEMANTIC MATCHES:
   - Look for semantic equivalence even when wording differs
   - Consider industry context and procurement terminology
   - Identify core procurement purpose beyond surface-level descriptions

You must return ONLY a JSON object with category_id and category_name."""

        user_prompt = f"""QUERY: Below are our categories retrieved by API call, so tell me for this Title of RFP "{title}", Description "{description}", and Current Category "{category}", which category best matches:

{categories_text}

Return ONLY the JSON object for the best matching category, like:
{{"category_id": 123, "category_name": "exact name from list"}}"""

        # Try each LLM service in order
        llm_services = [
            (self.gpt_client, "gpt-4o", "GitHub Copilot"),
            (self.groq_client, "mixtral-8x7b-32768", "Groq"),
            (
                self.openrouter_client,
                "meta-llama/llama-3.1-405b-instruct:free",
                "OpenRouter",
            ),
        ]

        last_error = None
        for client, model, service_name in llm_services:
            try:
                console.print(f"[cyan]Trying {service_name} LLM service...[/cyan]")

                response = client.chat.completions.create(
                    model=model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.1,
                    max_tokens=150,
                )

                result = self._parse_llm_response(response.choices[0].message.content)
                console.print(
                    f"[green]Successfully used {service_name} for matching[/green]"
                )
                return result, 1.0

            except Exception as e:
                last_error = e
                error_msg = str(e)
                if "429" in error_msg:  # Rate limit error
                    console.print(
                        f"[yellow]{service_name} rate limit reached, trying next service...[/yellow]"
                    )
                else:
                    console.print(
                        f"[yellow]{service_name} error: {error_msg}, trying next service...[/yellow]"
                    )
                continue

        # If all services fail
        logger.error(f"All LLM services failed. Last error: {str(last_error)}")
        return None, 0.0

    def match_by_majority(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using majority voting from similarity and LLM methods"""
        matches = []

        # Get matches from both methods
        methods = [(self.match_by_similarity, "Similarity"), (self.match_by_llm, "LLM")]

        for method_func, method_name in methods:
            try:
                match, score = method_func(title, description, category)
                if match:
                    matches.append((match, score, method_name))
                    console.print(
                        f"[cyan]{method_name}[/cyan] suggests: {match['category_name']} (score: {score:.2f})"
                    )
            except Exception as e:
                console.print(f"[yellow]{method_name} failed: {str(e)}[/yellow]")

        if not matches:
            return None, 0.0

        # Count occurrences of each category
        category_counts = {}
        for match, score, method in matches:
            cat_id = match["category_id"]
            if cat_id not in category_counts:
                category_counts[cat_id] = {
                    "count": 0,
                    "total_score": 0,
                    "category": match,
                    "methods": [],
                }
            category_counts[cat_id]["count"] += 1
            category_counts[cat_id]["total_score"] += score
            category_counts[cat_id]["methods"].append(method)

        # Find category with highest count and average score
        best_match = max(
            category_counts.values(),
            key=lambda x: (x["count"], x["total_score"] / x["count"]),
        )

        # Calculate confidence based on agreement and scores
        confidence = (best_match["count"] / len(matches)) * (
            best_match["total_score"] / best_match["count"]
        )

        console.print(
            f"\n[green]Majority vote selected:[/green] {best_match['category']['category_name']}"
        )
        console.print(
            f"[blue]Supporting methods:[/blue] {', '.join(best_match['methods'])}"
        )
        console.print(f"[blue]Confidence score:[/blue] {confidence:.2f}")

        return best_match["category"], confidence

    def _parse_llm_response(self, content: str) -> Dict:
        """Parse and validate LLM response"""
        try:
            # Clean the response
            content = content.strip()
            if not content.startswith("{"):
                # Extract JSON if wrapped in other text
                json_start = content.find("{")
                json_end = content.rfind("}") + 1
                if json_start >= 0 and json_end > json_start:
                    content = content[json_start:json_end]
                else:
                    raise ValueError("No JSON object found in response")

            # Parse JSON
            result = json.loads(content)

            # Validate required fields
            if "category_name" not in result or "category_id" not in result:
                raise ValueError("Missing required fields in response")

            # Find matching category from api_categories
            matched_category = next(
                (
                    cat
                    for cat in self.api_categories
                    if cat["category_id"] == result["category_id"]
                ),
                None,
            )

            if not matched_category:
                raise ValueError(
                    f"Category ID {result['category_id']} not found in API categories"
                )

            return matched_category

        except Exception as e:
            raise ValueError(f"Error parsing LLM response: {str(e)}")

    def _load_override_rules(self) -> Dict:
        """Load category override rules"""
        # These rules can be moved to a configuration file later
        return {
            "keywords": {
                "software": ["software", "license", "microsoft", "adobe", "digital"],
                "furniture": ["furniture", "chair", "desk", "table", "cabinet"],
                "medical": ["medical", "healthcare", "hospital", "clinical"],
                # Add more keyword-based rules
            },
            "exact_matches": {
                "Microsoft Office": "Software and Technology Products",
                "Office Supplies": "Office Supplies and Equipment",
                # Add more exact matches
            },
        }

    def _apply_override_rules(
        self, title: str, description: str, category: str
    ) -> Optional[Dict]:
        """Apply override rules to determine category"""
        # Convert inputs to lowercase for matching
        title_lower = title.lower()
        desc_lower = description.lower()
        category_lower = category.lower()

        # Check exact matches first
        for text, target_category in self.override_rules["exact_matches"].items():
            if text.lower() in title_lower or text.lower() in desc_lower:
                matching_category = next(
                    (
                        cat
                        for cat in self.api_categories
                        if cat["category_name"].lower() == target_category.lower()
                    ),
                    None,
                )
                if matching_category:
                    return matching_category

        # Check keyword-based rules
        keyword_matches = defaultdict(int)
        combined_text = f"{title_lower} {desc_lower} {category_lower}"

        for category_type, keywords in self.override_rules["keywords"].items():
            for keyword in keywords:
                if keyword in combined_text:
                    keyword_matches[category_type] += 1

        if keyword_matches:
            best_match = max(keyword_matches.items(), key=lambda x: x[1])
            if best_match[1] >= 2:  # Require at least 2 keyword matches
                # Find corresponding API category
                for cat in self.api_categories:
                    if best_match[0].lower() in cat["category_name"].lower():
                        return cat

        return None

    def match_by_enhanced_majority(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Enhanced majority voting with confidence thresholds and weights"""
        matches = []

        # First, check override rules
        override_match = self._apply_override_rules(title, description, category)
        if override_match:
            console.print("[cyan]Rule-based override applied[/cyan]")
            return override_match, 1.0

        # Get matches from both methods
        similarity_match, similarity_score = self.match_by_similarity(
            title, description, category
        )
        llm_match, llm_score = self.match_by_llm(title, description, category)

        # Apply confidence thresholds and weights
        weighted_matches = {}

        if similarity_match and similarity_score >= self.similarity_threshold:
            weighted_score = similarity_score * self.method_weights["similarity"]
            weighted_matches[similarity_match["category_id"]] = {
                "category": similarity_match,
                "score": weighted_score,
                "method": "similarity",
            }

        if llm_match:
            weighted_score = llm_score * self.method_weights["llm"]
            cat_id = llm_match["category_id"]
            if cat_id in weighted_matches:
                weighted_matches[cat_id]["score"] += weighted_score
            else:
                weighted_matches[cat_id] = {
                    "category": llm_match,
                    "score": weighted_score,
                    "method": "llm",
                }

        if not weighted_matches:
            return None, 0.0

        # Select best match based on weighted scores
        best_match = max(weighted_matches.values(), key=lambda x: x["score"])

        # Calculate final confidence score
        confidence = best_match["score"] / sum(self.method_weights.values())

        console.print(
            f"\n[green]Enhanced majority selected:[/green] {best_match['category']['category_name']}"
        )
        console.print(f"[blue]Method:[/blue] {best_match['method']}")
        console.print(f"[blue]Confidence score:[/blue] {confidence:.2f}")

        return best_match["category"], confidence

    def is_duplicate_bid(self, title: str, text2: str) -> bool:
        """Check if this bid has already been processed"""
        # Clean and normalize inputs
        title = " ".join(str(title).lower().split()).strip()
        text2 = " ".join(str(text2).lower().split()).strip()

        # Don't count as duplicate if title is empty
        if not title:
            return False

        # Generate bid identifier using title and either description or category
        bid_id = self._generate_bid_identifier(title, text2)

        # Check if it's a duplicate
        if bid_id in self.processed_bids:
            console.print(f"[yellow]Duplicate bid detected:[/yellow]")
            console.print(f"Title: {title[:100]}...")
            return True

        # Add to processed bids if not empty
        if bid_id.strip():
            self.processed_bids.add(bid_id)

        return False

    def _generate_bid_identifier(self, title: str, text2: str) -> str:
        """Generate a unique identifier for a bid based on its content"""
        # Normalize strings to handle minor differences
        title = " ".join(str(title).lower().split()).strip()
        text2 = " ".join(str(text2).lower().split()).strip()

        # Create identifier using title and either description or category
        title_part = title[:100] if title else ""
        text2_part = text2[:200] if text2 else ""

        return f"{title_part}::{text2_part}"

    def match_by_hybrid(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Enhanced hybrid matching with better accuracy"""
        try:
            # Clean and normalize inputs
            title = " ".join(title.lower().split()).strip()
            description = " ".join(description.lower().split()).strip()
            category = " ".join(category.lower().split()).strip()

            # Extract and clean category codes
            category_parts = []
            category_codes = []
            for part in category.split(";"):
                part = part.strip()
                if "-" in part:
                    code, desc = part.split("-", 1)
                    code = code.strip().rstrip("*")
                    desc = desc.strip().rstrip(";,.")
                    category_codes.append(code)
                    category_parts.extend([code, desc])

            # Domain-specific patterns and keywords
            domain_patterns = {
                "software": {
                    "primary": [
                        "software",
                        "system",
                        "application",
                        "platform",
                        "digital",
                        "it ",
                        "database",
                    ],
                    "secondary": [
                        "license",
                        "cloud",
                        "portal",
                        "program",
                        "web",
                        "online",
                        "computer",
                    ],
                    "codes": ["20", "92", "95"],
                },
                "construction": {
                    "primary": [
                        "construction",
                        "build",
                        "facility",
                        "infrastructure",
                        "renovation",
                    ],
                    "secondary": [
                        "installation",
                        "project",
                        "development",
                        "contractor",
                        "engineering",
                    ],
                    "codes": ["90", "91", "92"],
                },
                "services": {
                    "primary": [
                        "service",
                        "maintenance",
                        "support",
                        "consulting",
                        "professional",
                    ],
                    "secondary": [
                        "management",
                        "operation",
                        "provider",
                        "contractor",
                        "specialist",
                    ],
                    "codes": ["91", "92", "96"],
                },
                "equipment": {
                    "primary": ["equipment", "hardware", "device", "machine", "system"],
                    "secondary": [
                        "tool",
                        "apparatus",
                        "component",
                        "parts",
                        "supplies",
                    ],
                    "codes": ["03", "04", "07"],
                },
                "training": {
                    "primary": [
                        "training",
                        "education",
                        "learning",
                        "development",
                        "instruction",
                    ],
                    "secondary": [
                        "course",
                        "program",
                        "certification",
                        "workshop",
                        "class",
                    ],
                    "codes": ["92", "95", "96"],
                },
            }

            # Calculate base semantic similarity
            combined_text = f"{title} {title} {title} {category} {category} {description}"  # Title 3x, Category 2x
            text_embedding = self.model.encode([combined_text])[0]
            semantic_scores = cosine_similarity(
                [text_embedding], self.category_embeddings
            )[0]

            # Calculate enhanced fuzzy scores
            fuzzy_scores = np.zeros_like(semantic_scores)
            for idx, cat in enumerate(self.api_categories):
                cat_name = cat["category_name"].lower()

                # Multiple fuzzy matching algorithms for title
                title_scores = [
                    fuzz.token_set_ratio(title, cat_name) / 100.0,
                    fuzz.partial_ratio(title, cat_name) / 100.0,
                    fuzz.token_sort_ratio(title, cat_name) / 100.0,
                ]

                # Category code matching
                code_scores = []
                for code in category_codes:
                    if code in cat_name:
                        code_scores.append(1.0)  # Exact match
                    else:
                        code_scores.extend(
                            [
                                fuzz.ratio(code, cat_name[: len(code)])
                                / 100.0,  # Prefix match
                                fuzz.partial_ratio(code, cat_name)
                                / 100.0,  # Partial match
                            ]
                        )

                # Description matching if available
                desc_scores = []
                if description:
                    desc_scores = [
                        fuzz.token_set_ratio(description, cat_name) / 100.0,
                        fuzz.partial_ratio(description, cat_name) / 100.0,
                    ]

                # Combine scores with weights
                fuzzy_scores[idx] = (
                    0.5 * max(title_scores)  # Title most important
                    + 0.3 * (max(code_scores) if code_scores else 0.0)  # Code matching
                    + 0.2 * (max(desc_scores) if desc_scores else 0.0)  # Description
                )

            # Combine scores with adjusted weights
            final_scores = (0.65 * semantic_scores) + (0.35 * fuzzy_scores)

            # Apply domain-specific boosts
            for idx, score in enumerate(final_scores):
                cat_name = self.api_categories[idx]["category_name"].lower()
                boost = 1.0

                # Check each domain pattern
                for domain, patterns in domain_patterns.items():
                    # Primary keyword match
                    if any(kw in title.lower() for kw in patterns["primary"]):
                        if any(kw in cat_name for kw in patterns["primary"]):
                            boost *= 1.4

                    # Secondary keyword match
                    elif any(kw in title.lower() for kw in patterns["secondary"]):
                        if any(kw in cat_name for kw in patterns["secondary"]):
                            boost *= 1.2

                    # Code prefix match
                    if any(
                        code.startswith(prefix)
                        for code in category_codes
                        for prefix in patterns["codes"]
                    ):
                        if any(prefix in cat_name[:3] for prefix in patterns["codes"]):
                            boost *= 1.3

                # Apply the boost
                final_scores[idx] *= boost

            # Get best match
            best_idx = np.argmax(final_scores)
            best_score = final_scores[best_idx]

            # Calculate relative confidence
            score_mean = np.mean(final_scores)
            score_std = np.std(final_scores)
            relative_score = (
                (best_score - score_mean) / score_std if score_std > 0 else 0
            )

            # Dynamic threshold based on match quality
            threshold = 0.35  # Base threshold
            if relative_score > 2.0:  # Very strong match
                threshold = 0.3
            elif relative_score < 1.0:  # Weak match
                threshold = 0.4

            if best_score < threshold:
                return None, 0.0

            return self.api_categories[best_idx], float(best_score)

        except Exception as e:
            logger.error(f"Error in hybrid matching: {str(e)}")
            return None, 0.0

    def match_by_hierarchical(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using hierarchical approach with TF-IDF and code matching"""
        try:
            # Clean and normalize inputs
            title = " ".join(title.lower().split()).strip()
            description = " ".join(description.lower().split()).strip()
            category = " ".join(category.lower().split()).strip()

            # Extract category codes and descriptions
            category_codes = []
            category_descriptions = []
            for part in category.split(";"):
                part = part.strip()
                if "-" in part:
                    code, desc = part.split("-", 1)
                    code = code.strip().rstrip("*")
                    desc = desc.strip().rstrip(";,.")
                    category_codes.append(code)
                    category_descriptions.append(desc)

            # Create document corpus for TF-IDF
            corpus = [title]  # Start with title
            if description:
                corpus.append(description)  # Add description if available
            corpus.extend(category_descriptions)  # Add category descriptions

            # Add API category names to corpus
            api_categories_text = [
                cat["category_name"].lower() for cat in self.api_categories
            ]
            corpus.extend(api_categories_text)

            # Create TF-IDF matrix
            vectorizer = TfidfVectorizer(
                stop_words="english",
                ngram_range=(1, 2),  # Use both unigrams and bigrams
                max_features=1000,
            )
            tfidf_matrix = vectorizer.fit_transform(corpus)

            # Calculate similarities between input and API categories
            query_vector = tfidf_matrix[0:1]  # Use title as main query
            category_vectors = tfidf_matrix[-len(api_categories_text) :]
            similarities = cosine_similarity(query_vector, category_vectors)[0]

            # Get initial candidates (top 5)
            top_indices = np.argsort(similarities)[-5:][::-1]
            candidates = [(idx, similarities[idx]) for idx in top_indices]

            # Score adjustment based on category codes
            for idx, score in candidates:
                cat_name = self.api_categories[idx]["category_name"].lower()

                # Direct code match
                if any(code in cat_name for code in category_codes):
                    score *= 1.5

                # Code prefix match (first 2 digits)
                elif any(code[:2] in cat_name[:2] for code in category_codes):
                    score *= 1.3

                # Word overlap between category descriptions and API category
                cat_words = set(cat_name.split())
                desc_words = set(" ".join(category_descriptions).split())
                word_overlap = len(cat_words & desc_words)
                if word_overlap > 0:
                    score *= 1 + 0.1 * word_overlap

            # Get best match after adjustments
            best_idx, best_score = max(candidates, key=lambda x: x[1])

            # Calculate confidence score
            confidence = best_score
            if confidence < 0.3:  # Minimum confidence threshold
                return None, 0.0

            return self.api_categories[best_idx], float(confidence)

        except Exception as e:
            logger.error(f"Error in hierarchical matching: {str(e)}")
            return None, 0.0

    def match_by_original_similarity(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using original sentence embeddings method that achieved better accuracy"""
        try:
            # Combine input text with weights
            combined_text = (
                f"{title} {title} {description} {category}"  # Title weighted 2x
            )

            # Get embedding for input text
            text_embedding = self.model.encode([combined_text])[0]

            # Calculate similarities
            similarities = cosine_similarity(
                [text_embedding], self.category_embeddings
            )[0]

            # Get best match
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            return self.api_categories[best_idx], float(best_score)

        except Exception as e:
            logger.error(f"Error in original similarity matching: {str(e)}")
            return None, 0.0

    def match_by_weighted_fuzzy(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using weighted combination of fuzzy matching and semantic similarity"""
        try:
            # Clean and normalize inputs
            title = " ".join(title.lower().split()).strip()
            description = " ".join(description.lower().split()).strip()
            category = " ".join(category.lower().split()).strip()

            # Extract category codes and descriptions
            category_parts = []
            for part in category.split(";"):
                part = part.strip()
                if "-" in part:
                    code, desc = part.split("-", 1)
                    code = code.strip().rstrip("*")
                    desc = desc.strip().rstrip(";,.")
                    category_parts.extend([code, desc])

            # Get semantic similarity scores
            combined_text = (
                f"{title} {title} {description} {category}"  # Title weighted 2x
            )
            text_embedding = self.model.encode([combined_text])[0]
            semantic_scores = cosine_similarity(
                [text_embedding], self.category_embeddings
            )[0]

            # Calculate fuzzy match scores for each API category
            fuzzy_scores = np.zeros_like(semantic_scores)
            for idx, cat in enumerate(self.api_categories):
                cat_name = cat["category_name"].lower()

                # Title fuzzy match (highest weight)
                title_ratio = fuzz.token_set_ratio(title, cat_name) / 100.0

                # Category code match
                code_ratio = max(
                    (
                        fuzz.token_set_ratio(code, cat_name) / 100.0
                        for code in category_parts
                        if code.strip()
                    ),
                    default=0.0,
                )

                # Description match (if available)
                desc_ratio = (
                    fuzz.token_set_ratio(description, cat_name) / 100.0
                    if description
                    else 0.0
                )

                # Weighted combination of fuzzy scores
                fuzzy_scores[idx] = (
                    0.5 * title_ratio  # 50% weight to title
                    + 0.3 * code_ratio  # 30% weight to category codes
                    + 0.2 * desc_ratio  # 20% weight to description
                )

            # Combine semantic and fuzzy scores
            final_scores = (0.6 * semantic_scores) + (0.4 * fuzzy_scores)

            # Apply category-specific boosts
            for idx, score in enumerate(final_scores):
                cat_name = self.api_categories[idx]["category_name"].lower()

                # Boost for exact code matches
                if any(
                    code.strip() in cat_name for code in category_parts if code.strip()
                ):
                    final_scores[idx] *= 1.3

                # Boost for keyword matches
                keywords = {
                    "software": ["software", "system", "application", "digital"],
                    "construction": ["construction", "building", "facility"],
                    "services": ["service", "maintenance", "support"],
                    "equipment": ["equipment", "hardware", "device"],
                    "training": ["training", "education", "learning"],
                }

                for domain, domain_keywords in keywords.items():
                    if any(kw in cat_name for kw in domain_keywords):
                        if any(kw in title.lower() for kw in domain_keywords):
                            final_scores[idx] *= 1.2
                            break

            # Get best match
            best_idx = np.argmax(final_scores)
            best_score = final_scores[best_idx]

            # Apply confidence threshold
            if best_score < 0.35:  # Lower threshold since we're using multiple methods
                return None, 0.0

            return self.api_categories[best_idx], float(best_score)

        except Exception as e:
            logger.error(f"Error in weighted fuzzy matching: {str(e)}")
            return None, 0.0

    def match_by_ai_enhanced(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using combination of GPT-4 and weighted matching with fallback"""
        try:
            # First try GPT-4 matching
            llm_match = None
            llm_confidence = 0.0

            try:
                # Format categories for prompt
                categories_text = "\n".join(
                    [
                        f'{{"category_id": {cat["category_id"]}, "category_name": "{cat["category_name"]}"}}'
                        for cat in self.api_categories
                    ]
                )

                # Enhanced system prompt
                system_prompt = """You are an expert procurement data analyst specializing in government bid categorization. 
Your task is to match procurement requests to the most appropriate category from an authorized list.

CONTEXT:
You are part of a system that processes government bids and must categorize them accurately according to standardized categories.

CORE OBJECTIVE:
Analyze procurement details and select the MOST APPROPRIATE category from the provided list, considering:
1. Solicitation Title (Primary factor)
2. Current Category (Strong hint if available)
3. Description (Supporting context)

MATCHING RULES:
1. EXACT MATCHES:
   - If the title/description exactly matches a category name, prioritize that match
   - Consider industry-standard abbreviations and variations
   - Pay special attention to category codes in the current category

2. SEMANTIC MATCHES:
   - Look for semantic equivalence even when wording differs
   - Consider industry context and procurement terminology
   - Identify core procurement purpose beyond surface-level descriptions
   - Consider common government procurement patterns

You must return ONLY a JSON object with category_id and category_name."""

                # Enhanced user prompt
                user_prompt = f"""QUERY: Analyze this government bid and select the most appropriate category:

Title: "{title}"
Current Category: "{category}"
Description: "{description}"

Available categories:
{categories_text}

Consider:
1. Category codes in the current category (if any)
2. Industry-specific terminology
3. Common procurement patterns
4. Core purpose of the bid

Return ONLY the JSON object for the best matching category, like:
{{"category_id": 123, "category_name": "exact name from list"}}"""

                # Get GPT-4 completion
                response = self.gpt_client.chat.completions.create(
                    model="gpt-4o",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=0.1,
                    max_tokens=150,
                )

                llm_match = self._parse_llm_response(
                    response.choices[0].message.content
                )
                llm_confidence = 0.9  # High confidence for LLM match

            except Exception as e:
                logger.info(f"LLM matching failed, falling back to weighted: {str(e)}")

            # Get weighted matching result
            weighted_match, weighted_score = self.match_by_weighted_fuzzy(
                title, description, category
            )

            # If we have both matches, combine them
            if llm_match and weighted_match:
                if llm_match["category_id"] == weighted_match["category_id"]:
                    # Both methods agree - high confidence
                    return llm_match, max(llm_confidence, weighted_score)
                else:
                    # Methods disagree - use weighted scoring to decide
                    llm_weighted = 0.6  # Weight for LLM match
                    fuzzy_weighted = 0.4  # Weight for weighted fuzzy match

                    if llm_confidence * llm_weighted > weighted_score * fuzzy_weighted:
                        return llm_match, llm_confidence * llm_weighted
                    else:
                        return weighted_match, weighted_score * fuzzy_weighted

            # If only one method succeeded, use that
            if llm_match:
                return llm_match, llm_confidence
            if weighted_match:
                return weighted_match, weighted_score

            # If both methods failed
            return None, 0.0

        except Exception as e:
            logger.error(f"Error in AI-enhanced matching: {str(e)}")
            return None, 0.0

    def match_by_confident_hybrid(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Enhanced hybrid matching that only returns results for high-confidence matches"""
        try:
            # Clean and normalize inputs
            title = " ".join(title.lower().split()).strip()
            description = " ".join(description.lower().split()).strip()
            category = " ".join(category.lower().split()).strip()

            # Extract category codes and descriptions
            category_parts = []
            category_codes = []
            for part in category.split(";"):
                part = part.strip()
                if "-" in part:
                    code, desc = part.split("-", 1)
                    code = code.strip().rstrip("*")
                    desc = desc.strip().rstrip(";,.")
                    category_codes.append(code)
                    category_parts.extend([code, desc])

            # Expanded high-confidence category patterns
            high_confidence_patterns = {
                "software": {
                    "keywords": [
                        "software",
                        "license",
                        "microsoft",
                        "oracle",
                        "sap",
                        "adobe",
                        "digital",
                        "system",
                        "application",
                        "platform",
                        "cloud",
                        "saas",
                        "database",
                    ],
                    "codes": ["20", "208", "209", "43", "48"],
                },
                "construction": {
                    "keywords": [
                        "construction",
                        "building",
                        "renovation",
                        "facility",
                        "infrastructure",
                        "contractor",
                        "installation",
                    ],
                    "codes": ["90", "91", "236", "237", "238"],
                },
                "medical": {
                    "keywords": [
                        "medical",
                        "healthcare",
                        "hospital",
                        "clinical",
                        "health",
                        "pharmaceutical",
                        "medicine",
                        "dental",
                    ],
                    "codes": ["65", "339", "621", "622", "623"],
                },
                "professional_services": {
                    "keywords": [
                        "consulting",
                        "professional services",
                        "management services",
                        "advisory",
                        "consultant",
                    ],
                    "codes": ["91", "541", "611", "518", "519"],
                },
                "maintenance": {
                    "keywords": [
                        "maintenance",
                        "repair",
                        "servicing",
                        "support",
                        "upkeep",
                        "preventive",
                    ],
                    "codes": ["81", "811", "238", "561"],
                },
                "equipment": {
                    "keywords": [
                        "equipment",
                        "machinery",
                        "hardware",
                        "tools",
                        "apparatus",
                        "devices",
                    ],
                    "codes": ["23", "333", "334", "335"],
                },
                "supplies": {
                    "keywords": [
                        "supplies",
                        "materials",
                        "goods",
                        "products",
                        "items",
                        "consumables",
                    ],
                    "codes": ["42", "423", "424", "425"],
                },
                "training": {
                    "keywords": [
                        "training",
                        "education",
                        "learning",
                        "development",
                        "workshop",
                        "course",
                    ],
                    "codes": ["61", "611", "612"],
                },
            }

            # Calculate base semantic similarity with increased weight for title
            combined_text = f"{title} {title} {title} {title} {category} {category} {description}"  # Title weighted 4x
            text_embedding = self.model.encode([combined_text])[0]
            semantic_scores = cosine_similarity(
                [text_embedding], self.category_embeddings
            )[0]

            # Calculate enhanced fuzzy scores with adjusted weights
            fuzzy_scores = np.zeros_like(semantic_scores)
            for idx, cat in enumerate(self.api_categories):
                cat_name = cat["category_name"].lower()

                # Multiple fuzzy matching algorithms for title
                title_scores = [
                    fuzz.token_set_ratio(title, cat_name) / 100.0,
                    fuzz.partial_ratio(title, cat_name) / 100.0,
                    fuzz.token_sort_ratio(title, cat_name) / 100.0,
                ]

                # Code matching with higher weight
                code_scores = []
                for code in category_codes:
                    if code in cat_name:
                        code_scores.append(1.0)
                    else:
                        code_scores.extend(
                            [
                                fuzz.ratio(code, cat_name[: len(code)]) / 100.0,
                                fuzz.partial_ratio(code, cat_name) / 100.0,
                            ]
                        )

                # Description matching with adjusted weight
                desc_scores = []
                if description:
                    desc_scores = [
                        fuzz.token_set_ratio(description, cat_name) / 100.0,
                        fuzz.partial_ratio(description, cat_name) / 100.0,
                    ]

                # Combine scores with adjusted weights
                fuzzy_scores[idx] = (
                    0.6 * max(title_scores)  # Increased weight for title
                    + 0.25
                    * (
                        max(code_scores) if code_scores else 0.0
                    )  # Increased weight for codes
                    + 0.15
                    * (
                        max(desc_scores) if desc_scores else 0.0
                    )  # Reduced weight for description
                )

            # Combine scores with adjusted weights
            final_scores = (0.7 * semantic_scores) + (
                0.3 * fuzzy_scores
            )  # More weight to semantic matching

            # Apply confidence boosting with adjusted multipliers
            confidence_boost = 1.0
            for domain, patterns in high_confidence_patterns.items():
                # Check for keyword matches in title
                keyword_matches = sum(
                    1 for kw in patterns["keywords"] if kw in title.lower()
                )
                if keyword_matches > 0:
                    confidence_boost *= 1.0 + (
                        0.15 * keyword_matches
                    )  # Gradual boost based on matches

                # Check for code matches
                code_matches = sum(
                    1
                    for code in category_codes
                    for prefix in patterns["codes"]
                    if code.startswith(prefix)
                )
                if code_matches > 0:
                    confidence_boost *= 1.0 + (
                        0.1 * code_matches
                    )  # Gradual boost based on matches

            # Get best match
            best_idx = np.argmax(final_scores)
            best_score = final_scores[best_idx] * confidence_boost

            # Adjusted confidence thresholds
            base_threshold = 0.65  # Increased base threshold from 0.45

            # Additional confidence criteria with stricter conditions
            cat_name = self.api_categories[best_idx]["category_name"].lower()

            # Check for category code match (strict)
            code_match = any(
                code in cat_name or cat_name.split()[0].startswith(code)
                for code in category_codes
            )

            # Check for keyword match (strict)
            keyword_match = any(
                sum(
                    kw in title.lower() or kw in description.lower()
                    for kw in patterns["keywords"]
                )
                >= 2  # Require at least 2 keyword matches
                for patterns in high_confidence_patterns.values()
            )

            # Return match only if meets stricter criteria
            if best_score >= base_threshold and (
                code_match or keyword_match or best_score >= 0.8
            ):
                return self.api_categories[best_idx], float(best_score)

            # Return None for low-confidence matches
            return None, 0.0

        except Exception as e:
            logger.error(f"Error in confident hybrid matching: {str(e)}")
            return None, 0.0

    def match_by_fuzzy_ollama(
        self, title: str, description: str, category: str
    ) -> Tuple[Dict, float]:
        """Match using combination of Ollama LLM and hybrid matching"""
        try:
            # First try Ollama matching
            ollama_match = None
            ollama_confidence = 0.0

            try:
                # Use llama3.2:3b model explicitly
                model_name = "llama3.2:3b"
                console.print(f"[cyan]Using Ollama model: {model_name}[/cyan]")

                # Format categories for prompt
                categories_text = "\n".join(
                    [
                        f'{{"category_id": {cat["category_id"]}, "category_name": "{cat["category_name"]}"}}'
                        for cat in self.api_categories
                    ]
                )

                # Enhanced system prompt for Ollama
                system_prompt = """You are an expert procurement data analyst. Your task is to match government bids to the most appropriate category.

RULES:
1. Analyze the bid's Title (most important), Category, and Description
2. Match to the most appropriate category from the provided list
3. Consider industry terminology and procurement patterns
4. Return ONLY a JSON object with category_id and category_name
5. Focus on exact matches first, then semantic matches

IMPORTANT: Your response must be a valid JSON object in this exact format:
{"category_id": number, "category_name": "string"}"""

                # Concise user prompt for Ollama
                user_prompt = f"""Match this bid to a category:
Title: "{title}"
Category: "{category}"
Description: "{description}"

Available categories:
{categories_text}

Return ONLY a JSON object like: {{"category_id": ID, "category_name": "EXACT NAME"}}"""

                # Try Ollama with llama3.2:3b model
                try:
                    console.print("[cyan]Trying Ollama LLM service...[/cyan]")
                    response = self.ollama_client.chat.completions.create(
                        model=model_name,
                        messages=[
                            {"role": "system", "content": system_prompt},
                            {"role": "user", "content": user_prompt},
                        ],
                        temperature=0.1,
                        max_tokens=150,
                    )

                    # Extract content and clean it
                    content = response.choices[0].message.content.strip()
                    # Find JSON in the response
                    json_start = content.find("{")
                    json_end = content.rfind("}") + 1
                    if json_start >= 0 and json_end > json_start:
                        json_str = content[json_start:json_end]
                        try:
                            result = json.loads(json_str)
                            if "category_id" in result and "category_name" in result:
                                ollama_match = next(
                                    (
                                        cat
                                        for cat in self.api_categories
                                        if cat["category_id"] == result["category_id"]
                                    ),
                                    None,
                                )
                                if ollama_match:
                                    ollama_confidence = 0.85
                                    console.print(
                                        "[green]Successfully used Ollama for matching[/green]"
                                    )
                                else:
                                    raise ValueError(
                                        f"Category ID {result['category_id']} not found"
                                    )
                            else:
                                raise ValueError("Missing required fields in response")
                        except json.JSONDecodeError:
                            raise ValueError("Invalid JSON format in response")
                    else:
                        raise ValueError("No JSON object found in response")

                except Exception as e:
                    console.print(
                        f"[yellow]Ollama error: {str(e)}, falling back to hybrid[/yellow]"
                    )

            except Exception as e:
                logger.info(f"Ollama matching failed, using hybrid: {str(e)}")

            # Get hybrid matching result as fallback
            hybrid_match, hybrid_score = self.match_by_hybrid(
                title, description, category
            )

            # If we have both matches, combine them
            if ollama_match and hybrid_match:
                if ollama_match["category_id"] == hybrid_match["category_id"]:
                    # Both methods agree - high confidence
                    return ollama_match, max(ollama_confidence, hybrid_score)
                else:
                    # Methods disagree - use weighted scoring
                    ollama_weighted = 0.5  # Equal weight for Ollama
                    hybrid_weighted = 0.5  # Equal weight for hybrid

                    if (
                        ollama_confidence * ollama_weighted
                        > hybrid_score * hybrid_weighted
                    ):
                        return ollama_match, ollama_confidence * ollama_weighted
                    else:
                        return hybrid_match, hybrid_score * hybrid_weighted

            # If only one method succeeded, use that
            if ollama_match:
                return ollama_match, ollama_confidence
            if hybrid_match:
                return hybrid_match, hybrid_score

            # If both methods failed
            return None, 0.0

        except Exception as e:
            logger.error(f"Error in fuzzy-ollama matching: {str(e)}")
            return None, 0.0
