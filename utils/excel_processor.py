import pandas as pd
import requests
from sentence_transformers import SentenceTransformer
from fuzzywuzzy import fuzz
import spacy
import logging
from typing import Dict, List, Tuple, Optional, Any
import numpy as np
from pathlib import Path
import urllib3
import warnings
import json
import os
import re
from openai import OpenAI
from langchain_groq import ChatGroq
from langchain.schema import HumanMessage, SystemMessage
from time import sleep
from rich.console import Console
from azure.ai.inference import ChatCompletionsClient
from azure.ai.inference.models import SystemMessage as AzureSystemMessage
from azure.ai.inference.models import UserMessage as AzureUserMessage
from azure.core.credentials import AzureKeyCredential
from mistralai import (
    Mistral,
    UserMessage as MistralUserMessage,
    SystemMessage as MistralSystemMessage,
)
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
from cachetools import TTLCache, cached
from functools import lru_cache
import concurrent.futures
from typing import Dict, List, Tuple, Optional, Set
from collections import defaultdict

# Suppress SSL verification warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# Configure logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

console = Console()

# Cache configurations
API_CACHE_TTL = 3600  # 1 hour TTL for API responses
API_CACHE_SIZE = 1000  # Maximum number of cached items
EMBEDDING_CACHE_SIZE = 10000  # Maximum number of cached embeddings

class ExcelProcessor:
    def __init__(self):
        try:
            # Disable progress bars for sentence transformers
            logging.getLogger("sentence_transformers").setLevel(logging.WARNING)

            self.sentence_model = SentenceTransformer("paraphrase-MiniLM-L6-v2")
            self.nlp = spacy.load("en_core_web_sm")
            
            # Initialize caches
            self.api_cache = TTLCache(maxsize=API_CACHE_SIZE, ttl=API_CACHE_TTL)
            self.embedding_cache = TTLCache(maxsize=EMBEDDING_CACHE_SIZE, ttl=API_CACHE_TTL)
            self.processed_bids = set()  # For tracking processed bids
            
            # Efficient data structures for category matching
            self.category_embeddings = None
            self.category_keywords = defaultdict(set)
            self.category_codes = defaultdict(set)
            
        except Exception as e:
            logger.error(f"Error loading AI models: {str(e)}")
            raise

        # API endpoints
        self.API_ENDPOINTS = {
            "notice": "https://bidsportal.com/api/getNotice",
            "category": "https://bidsportal.com/api/getCategory",
            "agency": "https://bidsportal.com/api/getAgency",
            "state": "https://bidsportal.com/api/getState",
        }

        self.console = Console()

    @cached(cache=TTLCache(maxsize=API_CACHE_SIZE, ttl=API_CACHE_TTL))
    def fetch_api_data_cached(self, endpoint_key: str, params_str: str = "") -> List[Dict]:
        """Cached version of fetch_api_data"""
        params = json.loads(params_str) if params_str else None
        return self.fetch_api_data(endpoint_key, params)

    @lru_cache(maxsize=EMBEDDING_CACHE_SIZE)
    def get_embedding_cached(self, text: str) -> np.ndarray:
        """Cached version of sentence embedding computation"""
        return self.sentence_model.encode([text])[0]

    def process_dataframe_parallel(self, df: pd.DataFrame, categories: List[Dict], 
                                 notice_types: List[Dict], agencies: List[Dict], 
                                 states: List[Dict]) -> pd.DataFrame:
        """Process DataFrame in parallel using vectorized operations where possible"""
        # Pre-allocate new columns with default values
        df['API Category'] = None
        df['API Category ID'] = None
        df['API Notice Type'] = None
        df['API Agency'] = None
        df['API State'] = None

        # Vectorized operations for text preprocessing
        df['processed_title'] = df['Title'].fillna('').astype(str).str.lower().str.strip()
        df['processed_desc'] = df['Description'].fillna('').astype(str).str.lower().str.strip()
        df['processed_category'] = df['Category'].fillna('').astype(str).str.lower().str.strip()

        # Process chunks in parallel
        chunk_size = max(1, len(df) // (os.cpu_count() or 1))
        chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

        def process_chunk(chunk_df: pd.DataFrame) -> pd.DataFrame:
            """Process a chunk of the DataFrame using vectorized operations"""
            try:
                # Create result DataFrame with same index as chunk
                result_df = pd.DataFrame(index=chunk_df.index)
                
                # Apply category matching using pandas apply
                category_results = chunk_df.apply(
                    lambda row: self.find_best_category_match(
                        row['processed_title'],
                        row['processed_desc'],
                        row['processed_category'],
                        categories
                    ),
                    axis=1,
                    result_type='expand'
                )
                result_df['API Category'] = category_results.apply(lambda x: x[0] if x is not None else None)
                result_df['API Category ID'] = category_results.apply(lambda x: x[1] if x is not None else None)

                # Apply notice type matching
                notice_results = chunk_df.apply(
                    lambda row: self.determine_notice_type(
                        f"{row['processed_title']} {row['processed_desc']}", 
                        notice_types
                    ),
                    axis=1,
                    result_type='expand'
                )
                result_df['API Notice Type'] = notice_results.apply(lambda x: x[0] if x is not None else None)

                # Apply agency matching
                agency_results = chunk_df.apply(
                    lambda row: self.find_best_agency_match(
                        row.get('Agency', ''),
                        row.get('Bid Detail Page URL', ''),
                        agencies
                    ),
                    axis=1,
                    result_type='expand'
                )
                result_df['API Agency'] = agency_results.apply(lambda x: x[0] if x is not None else None)

                # Apply state matching
                state_results = chunk_df.apply(
                    lambda row: self.find_state_match(
                        row['processed_desc'],
                        row.get('Agency', ''),
                        row.get('Bid Detail Page URL', ''),
                        states
                    ),
                    axis=1,
                    result_type='expand'
                )
                result_df['API State'] = state_results.apply(lambda x: x[0] if x is not None else None)

                return result_df

            except Exception as e:
                logger.error(f"Error processing chunk: {str(e)}")
                # Return empty DataFrame with correct columns on error
                return pd.DataFrame(
                    {
                        'API Category': None,
                        'API Category ID': None,
                        'API Notice Type': None,
                        'API Agency': None,
                        'API State': None
                    },
                    index=chunk_df.index
                )

        # Process chunks in parallel
        with ThreadPoolExecutor(max_workers=os.cpu_count() or 1) as executor:
            futures = [executor.submit(process_chunk, chunk) for chunk in chunks]
            results_dfs = []
            for future in as_completed(futures):
                try:
                    result_df = future.result()
                    results_dfs.append(result_df)
                except Exception as e:
                    logger.error(f"Error processing chunk: {str(e)}")

        # Combine results
        if results_dfs:
            results_df = pd.concat(results_dfs)
            # Update original DataFrame with results
            for col in ['API Category', 'API Category ID', 'API Notice Type', 'API Agency', 'API State']:
                df[col] = results_df[col]

        # Clean up temporary columns
        df.drop(['processed_title', 'processed_desc', 'processed_category'], axis=1, inplace=True)

        return df

    def is_processing_complete(self, df: pd.DataFrame) -> bool:
        """Check if all required columns have been processed"""
        required_columns = ['API Category', 'API Category ID', 'API Notice Type', 'API Agency', 'API State']
        
        # Check if all required columns exist
        if not all(col in df.columns for col in required_columns):
            return False
            
        # Check if any required column has all None values
        if any(df[col].isna().all() for col in required_columns):
            return False
            
        # Check if processing is still ongoing (more than 50% None values in any column)
        if any(df[col].isna().mean() > 0.5 for col in required_columns):
            return False
            
        return True

    def process_excel_file(self, excel_path: str) -> bool:
        """Process Excel file with optimized parallel processing and caching"""
        try:
            if not os.path.exists(excel_path):
                logger.error(f"Excel file not found: {excel_path}")
                return False

            # Read Excel file efficiently
            df = pd.read_excel(excel_path)
            logger.info(f"Processing Excel file: {excel_path}")
            print(f"\nProcessing Excel file: {excel_path}")

            # Fetch API data with caching
            print("\nFetching data from APIs...")
            categories = self.fetch_api_data_cached("category")
            notice_types = self.fetch_api_data_cached("notice")
            agencies = self.fetch_api_data_cached("agency")
            states = self.fetch_api_data_cached("state", json.dumps({"country_id": 10}))

            # Process DataFrame in parallel with optimizations
            df = self.process_dataframe_parallel(df, categories, notice_types, agencies, states)

            # Check if processing is complete
            if not self.is_processing_complete(df):
                logger.error(f"Excel processing incomplete for {excel_path}")
                print(f"❌ Excel processing incomplete for {excel_path}")
                return False

            # Save processed file
            output_path = str(Path(excel_path).with_name(f"{Path(excel_path).name}"))
            df.to_excel(output_path, index=False)
            print(f"\n✅ Successfully saved processed file to: {output_path}")
            
            return True

        except Exception as e:
            logger.error(f"Error processing Excel file: {str(e)}")
            print(f"❌ Error processing Excel file: {str(e)}")
            return False

    def make_request(self, url: str, params: Dict = None) -> requests.Response:
        """Make HTTP request with proper error handling and SSL verification disabled"""
        try:
            response = requests.post(
                url,
                json=params or {},
                verify=False,  # Disable SSL verification
                timeout=30,  # Set timeout to 30 seconds
                headers={
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
                },
            )
            response.raise_for_status()
            return response
        except requests.exceptions.SSLError as e:
            logger.warning(f"SSL Error: {str(e)}")
            logger.info("Attempting request without SSL verification...")
            return requests.post(url, json=params or {}, verify=False)
        except requests.exceptions.ConnectionError as e:
            logger.error(f"Connection Error: {str(e)}")
            raise
        except requests.exceptions.Timeout as e:
            logger.error(f"Timeout Error: {str(e)}")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"Request Error: {str(e)}")
            raise

    def fetch_api_data(self, endpoint_key: str, params: Dict = None) -> List[Dict]:
        """Fetch data from API with caching and proper error handling"""
        if endpoint_key in self.api_cache:
            return self.api_cache[endpoint_key]

        try:
            url = self.API_ENDPOINTS[endpoint_key]
            response = self.make_request(url, params)

            try:
                response_data = response.json()

                # All responses are nested under 'data' key
                if isinstance(response_data, dict) and "data" in response_data:
                    data = response_data["data"]
                else:
                    logger.warning(
                        f"Unexpected response format from {endpoint_key} API - missing 'data' key"
                    )
                    return []

                # Transform data based on endpoint
                transformed_data = []

                if endpoint_key == "category":
                    for item in data:
                        if isinstance(item, dict):
                            transformed_data.append(
                                {
                                    "name": item.get("category_name", ""),
                                    "id": item.get("category_id"),
                                    "raw_data": item,  # Store complete raw data
                                }
                            )

                elif endpoint_key == "agency":
                    for item in data:
                        if isinstance(item, dict):
                            transformed_data.append(
                                {
                                    "name": item.get("agency_name", ""),
                                    "id": item.get("agency_id"),
                                    "code": item.get("agency_code"),
                                    "raw_data": item,
                                }
                            )

                elif endpoint_key == "state":
                    for item in data:
                        if isinstance(item, dict):
                            transformed_data.append(
                                {
                                    "name": item.get("state_name", ""),
                                    "id": item.get("state_id"),
                                    "code": item.get("state_code"),
                                    "country_code": item.get("state_country_code"),
                                    "country_id": item.get("state_country_id"),
                                    "country_name": item.get("state_country_name"),
                                    "raw_data": item,
                                }
                            )

                elif endpoint_key == "notice":
                    for item in data:
                        if isinstance(item, dict):
                            transformed_data.append(
                                {
                                    "name": item.get("notice_type", ""),
                                    "id": item.get("notice_id"),
                                    "sort": item.get("sort"),
                                    "background_color": item.get(
                                        "backround_color"
                                    ),  # Note: API has typo in field name
                                    "raw_data": item,
                                }
                            )

                # Cache the transformed data
                self.api_cache[endpoint_key] = transformed_data

                # Log results
                logger.info(
                    f"Retrieved {len(transformed_data)} items from {endpoint_key} API"
                )
                print(
                    f"Retrieved {len(transformed_data)} items from {endpoint_key} API"
                )

                if transformed_data:
                    logger.debug(f"Sample {endpoint_key} data: {transformed_data[0]}")
                    print(f"Sample {endpoint_key} data: {transformed_data[0]}")

                return transformed_data

            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON response from {endpoint_key} API: {str(e)}")
                return []

        except Exception as e:
            logger.error(f"Error fetching {endpoint_key} data: {str(e)}")
            return []

    def get_embedding_similarity(self, text1: str, text2: str) -> float:
        """Calculate semantic similarity between two texts"""
        try:
            # Convert inputs to strings and handle None/nan values
            text1 = str(text1).strip() if pd.notna(text1) else ""
            text2 = str(text2).strip() if pd.notna(text2) else ""

            if not text1 or not text2:
                return 0.0

            embedding1 = self.sentence_model.encode([text1])[0]
            embedding2 = self.sentence_model.encode([text2])[0]
            return np.dot(embedding1, embedding2) / (
                np.linalg.norm(embedding1) * np.linalg.norm(embedding2)
            )
        except Exception as e:
            logger.error(f"Error calculating embedding similarity: {str(e)}")
            return 0.0

    def _prepare_embeddings(self):
        """Prepare embeddings for all API categories with caching"""
        try:
            category_texts = [cat["category_name"] for cat in self.api_categories]
            
            # Use cached embeddings where possible
            self.category_embeddings = np.array([
                self.get_embedding_cached(text) for text in category_texts
            ])
            
            # Build efficient lookup structures
            for cat in self.api_categories:
                cat_name = cat["category_name"].lower()
                # Extract keywords
                words = set(cat_name.split())
                self.category_keywords[cat["id"]].update(words)
                
                # Extract codes if present (assuming format like "123-Description")
                if "-" in cat_name:
                    code = cat_name.split("-")[0].strip()
                    self.category_codes[cat["id"]].add(code)
        except Exception as e:
            logger.error(f"Error preparing embeddings: {str(e)}")
            raise

    @cached(cache=TTLCache(maxsize=1000, ttl=3600))
    def find_best_category_match_cached(self, title: str, description: str, 
                                      current_category: str, categories_json: str) -> Tuple[Optional[str], Optional[int]]:
        """Cached version of category matching"""
        try:
            categories = json.loads(categories_json)
            return self.find_best_category_match(title, description, current_category, categories)
        except Exception as e:
            logger.error(f"Error in cached category matching: {str(e)}")
            return None, None

    def find_best_category_match(self, title: str, description: str, 
                                current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Optimized category matching using cached embeddings and efficient data structures"""
        try:
            # Clean and normalize inputs
            title = str(title).strip().lower() if pd.notna(title) else ""
            description = str(description).strip().lower() if pd.notna(description) else ""
            current_category = str(current_category).strip().lower() if pd.notna(current_category) else ""

            if not title and not description:
                return None, None

            # Generate bid identifier for duplicate detection
            bid_id = self._generate_bid_identifier(title, description)
            if bid_id in self.processed_bids:
                logger.info(f"Duplicate bid detected: {title[:50]}...")
                return None, None
            self.processed_bids.add(bid_id)

            # Get cached embeddings for input text
            combined_text = f"{title} {title} {description} {current_category}"  # Weight title more heavily
            text_embedding = self.get_embedding_cached(combined_text)

            # Calculate similarities using vectorized operations
            similarities = np.dot(self.category_embeddings, text_embedding) / (
                np.linalg.norm(self.category_embeddings, axis=1) * np.linalg.norm(text_embedding)
            )

            # Get top candidates
            top_k = 5
            top_indices = np.argpartition(similarities, -top_k)[-top_k:]
            candidates = [(idx, similarities[idx]) for idx in top_indices]

            # Score adjustment using efficient data structures
            for idx, score in candidates:
                cat = self.api_categories[idx]
                cat_id = cat["id"]
                
                # Check keyword matches
                text_words = set(combined_text.split())
                keyword_matches = len(text_words & self.category_keywords[cat_id])
                if keyword_matches > 0:
                    similarities[idx] *= 1 + (0.1 * keyword_matches)

                # Check code matches
                if self.category_codes[cat_id]:
                    if any(code in combined_text for code in self.category_codes[cat_id]):
                        similarities[idx] *= 1.3
                    elif any(code[:2] in combined_text[:2] for code in self.category_codes[cat_id]):
                        similarities[idx] *= 1.2

            # Get best match after adjustments
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            # Apply confidence threshold
            if best_score < 0.3:
                return None, None

            return self.api_categories[best_idx]["name"], self.api_categories[best_idx]["id"]

        except Exception as e:
            logger.error(f"Error in category matching: {str(e)}")
            return None, None

    def _generate_bid_identifier(self, title: str, description: str) -> str:
        """Generate unique identifier for bid using efficient string operations"""
        # Normalize strings
        title = "".join(c.lower() for c in str(title) if c.isalnum())
        desc = "".join(c.lower() for c in str(description) if c.isalnum())
        
        # Create compact identifier
        return f"{title[:100]}::{desc[:100]}"

    def _get_system_prompt(self) -> str:
        """Get the enhanced system prompt for category matching"""
        return """You are an expert procurement data analyst specializing in government bid categorization. Your task is to match procurement requests to the most appropriate category from an authorized list.

CONTEXT:
You are part of a system that processes government bids and must categorize them accurately according to standardized categories. Your categorization will be used for bid management and AWS upload.

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

3. HIERARCHICAL MATCHING:
   - If uncertain between specific and general categories:
     * Choose specific when confidence is high
     * Default to general when uncertain
   - Consider parent-child relationships in categories

4. CONFIDENCE REQUIREMENTS:
   - Must be highly confident in category assignment
   - When in doubt between multiple categories:
     * Prioritize based on title relevance
     * Consider current category as strong signal
     * Use description for disambiguation

RESPONSE FORMAT:
You must respond ONLY with a JSON object in this exact format:
{
    "category_id": number,
    "category_name": "exact_name_from_list"
}

CRITICAL REQUIREMENTS:
- Must use EXACT category names from provided list
- Must include valid category ID
- Must return valid JSON only
- NO explanations or additional text
- NO partial matches or modifications to category names
- NO creating new categories

Your role is crucial in ensuring accurate bid categorization for government procurement systems."""

    def _format_analysis_prompt(
        self,
        title: str,
        description: str,
        current_category: str,
        model_responses: List[Dict],
        categories_json: str,
    ) -> str:
        """Format the prompt for analyzing model responses"""
        responses_text = "\n".join(
            [
                f"{resp['model']}: {json.dumps(resp['response'])}"
                for resp in model_responses
            ]
        )

        return f"""Analyze these model suggestions and determine the most appropriate category.

REQUEST DETAILS:
Title: {title}
Current Category: {current_category}
Description: {description}

MODEL SUGGESTIONS:
{responses_text}

AVAILABLE CATEGORIES:
{categories_json}

INSTRUCTIONS:
1. Consider all model suggestions
2. Evaluate the confidence and reasoning behind each suggestion
3. Consider the current category as a strong hint
4. When models disagree, choose the most appropriate category based on:
   - Relevance to the procurement request
   - Specificity vs. generality
   - Industry standard categorization
5. Return ONLY a JSON object with the final category_id and category_name

IMPORTANT:
- Must return valid JSON
- Must use exact category names from the list
- Must include both id and name
- No explanations or additional text"""

    def _get_majority_vote(
        self, model_responses: List[Dict]
    ) -> Tuple[Optional[str], Optional[int]]:
        """Get majority vote from model responses as fallback"""
        if not model_responses:
            return None, None

        # Count category occurrences
        category_counts = {}
        for resp in model_responses:
            response = resp["response"]
            key = (response["category_id"], response["category_name"])
            category_counts[key] = category_counts.get(key, 0) + 1

        # Get most common category
        most_common = max(category_counts.items(), key=lambda x: x[1])
        category = most_common[0]

        return category[1], category[0]

    def _format_category_prompt(
        self, title: str, description: str, current_category: str, categories_json: str
    ) -> str:
        """Format the enhanced category matching prompt"""
        return f"""Analyze this government procurement request and determine the most appropriate category.

PROCUREMENT DETAILS:
Title: {title}
Current Category: {current_category}
Description: {description}

AVAILABLE CATEGORIES:
{categories_json}

MATCHING INSTRUCTIONS:
1. PRIMARY ANALYSIS:
   - Analyze title for key procurement focus
   - Identify industry sector and procurement type
   - Note specific products/services mentioned

2. CATEGORY VALIDATION:
   - Compare against provided category list
   - Consider current category as strong signal
   - Verify against description details

3. DECISION PROCESS:
   - Match core procurement purpose to categories
   - Consider industry standard classifications
   - Evaluate specificity vs. generality
   - Ensure regulatory compliance

4. QUALITY CHECKS:
   - Verify category exists in provided list
   - Confirm ID matches category name
   - Ensure exact name match from list

RESPONSE REQUIREMENTS:
- Must return valid JSON object
- Must use exact category names
- Must include correct category ID
- No additional text or explanations

Example valid response:
{{"category_id": 123, "category_name": "Information Technology"}}"""

    def determine_notice_type(
        self, text: str, notice_types: List[Dict]
    ) -> Tuple[Optional[str], Optional[int]]:
        """Determine the notice type based on text analysis"""
        try:
            if not text or not notice_types:
                return None, None

            text = text.lower()

            # Get valid notice types from API
            api_notice_types = {
                nt["name"].lower(): (nt["name"], nt["id"]) for nt in notice_types
            }

            # Mapping according to excel_extra_cols.md
            notice_mappings = {
                "rfp": "Request For Proposal",
                "rfq": "Request For Proposal",
                "rfx": "Request For Proposal",
                "invitation for bid": "Request For Proposal",
                "request for proposal": "Request For Proposal",
                "request for quote": "Request For Proposal",
                "bid invitation": "Request For Proposal",
                "sources sought": "Sources Sought / RFI",
                "rfi": "Sources Sought / RFI",
                "award": "Award Notice",
            }

            # First check exact matches with API notice types
            for api_type, (original_name, api_id) in api_notice_types.items():
                if api_type in text:
                    return original_name, api_id

            # Then check mapped variations
            for key, mapped_type in notice_mappings.items():
                if key in text and mapped_type.lower() in api_notice_types:
                    return api_notice_types[mapped_type.lower()]

            # If no match found but text contains "bid" or "solicitation", default to RFP
            if any(word in text for word in ["bid", "solicitation"]):
                return api_notice_types["request for proposal"]

            return None, None

        except Exception as e:
            logger.error(f"Error determining notice type: {str(e)}")
            return None, None

    def find_best_agency_match(
        self, agency_name: str, bid_url: str, agencies: List[Dict]
    ) -> Tuple[Optional[str], Optional[str]]:
        """Find the best matching agency using AI and URL analysis"""
        try:
            # Convert inputs to strings and handle None/nan values
            agency_name = str(agency_name).strip() if pd.notna(agency_name) else ""
            bid_url = str(bid_url).strip() if pd.notna(bid_url) else ""

            if not agency_name and not bid_url:
                return None, None

            best_match = None
            best_score = 0
            best_id = None

            # Combine agency name and URL for matching
            combined_text = f"{agency_name} {bid_url}"

            for agency in agencies:
                if (
                    not isinstance(agency, dict)
                    or "name" not in agency
                    or "id" not in agency
                ):
                    continue

                agency_name = (
                    str(agency["name"]).strip() if pd.notna(agency["name"]) else ""
                )
                if not agency_name:
                    continue

                # Calculate similarity scores
                semantic_score = self.get_embedding_similarity(
                    combined_text, agency_name
                )
                fuzzy_score = (
                    fuzz.token_sort_ratio(combined_text.lower(), agency_name.lower())
                    / 100
                )

                # Combine scores with weights
                combined_score = (semantic_score * 0.7) + (fuzzy_score * 0.3)

                if combined_score > best_score and combined_score > 0.6:
                    best_score = combined_score
                    best_match = agency_name
                    best_id = agency["id"]

            return best_match, best_id

        except Exception as e:
            logger.error(f"Error finding agency match: {str(e)}")
            return None, None

    def find_state_match(
        self, description: str, agency_name: str, bid_url: str, states: List[Dict]
    ) -> Tuple[Optional[str], Optional[str]]:
        """Find the state match using multiple data points"""
        try:
            # Convert inputs to strings and handle None/nan values
            description = str(description).strip() if pd.notna(description) else ""
            agency_name = str(agency_name).strip() if pd.notna(agency_name) else ""
            bid_url = str(bid_url).strip() if pd.notna(bid_url) else ""

            if not any([description, agency_name, bid_url]):
                return None, None

            # Combine all text for analysis
            combined_text = f"{description} {agency_name} {bid_url}"

            # Extract state names using spaCy
            doc = self.nlp(combined_text)
            potential_states = [ent.text for ent in doc.ents if ent.label_ == "GPE"]

            best_match = None
            best_score = 0
            best_id = None

            for state in states:
                if (
                    not isinstance(state, dict)
                    or "name" not in state
                    or "id" not in state
                ):
                    continue

                state_name = (
                    str(state["name"]).strip() if pd.notna(state["name"]) else ""
                )
                if not state_name:
                    continue

                # Check each potential state against API states
                for potential_state in potential_states:
                    potential_state = str(potential_state).strip()
                    semantic_score = self.get_embedding_similarity(
                        potential_state, state_name
                    )
                    fuzzy_score = (
                        fuzz.token_sort_ratio(
                            potential_state.lower(), state_name.lower()
                        )
                        / 100
                    )

                    combined_score = (semantic_score * 0.7) + (fuzzy_score * 0.3)

                    if combined_score > best_score and combined_score > 0.6:
                        best_score = combined_score
                        best_match = state_name
                        best_id = state["id"]

            return best_match, best_id

        except Exception as e:
            logger.error(f"Error finding state match: {str(e)}")
            return None, None

    def process_completed_folder(self, folder_path: str) -> bool:
        """Process all Excel files in a completed folder"""
        try:
            # Find all Excel files in the folder
            excel_files = list(Path(folder_path).glob("*.xlsx"))

            if not excel_files:
                logger.warning(f"No Excel files found in {folder_path}")
                return False

            success = True
            for excel_file in excel_files:
                logger.info(f"Processing Excel file: {excel_file}")
                print(f"\n📊 Processing Excel file: {excel_file}")

                if not self.process_excel_file(str(excel_file)):
                    success = False
                    logger.error(f"Failed to process {excel_file}")
                    print(f"❌ Failed to process {excel_file}")
                else:
                    print(f"✅ Successfully processed {excel_file}")

            return success

        except Exception as e:
            logger.error(f"Error processing folder {folder_path}: {str(e)}")
            print(f"❌ Error processing folder: {str(e)}")
            return False

    def _format_enhanced_analysis_prompt(
        self,
        title: str,
        description: str,
        current_category: str,
        model_responses: List[Dict],
        categories_json: str,
        confidence: float,
    ) -> str:
        """Format an enhanced analysis prompt that includes confidence information"""
        responses_text = "\n".join(
            [
                f"{resp['model']}: {json.dumps(resp['response'])}"
                for resp in model_responses
            ]
        )

        return f"""Analyze these model suggestions and determine the most appropriate category.
Current model agreement confidence: {confidence:.1%}

PROCUREMENT DETAILS:
Title: {title}
Current Category: {current_category}
Description: {description}

MODEL SUGGESTIONS:
{responses_text}

AVAILABLE CATEGORIES:
{categories_json}

ANALYSIS REQUIREMENTS:
1. Evaluate each model's suggestion considering:
   - Semantic relevance to procurement details
   - Industry standard categorization practices
   - Hierarchical category relationships
   
2. Consider confidence level:
   - High confidence ({confidence >= 0.7}): Validate majority opinion
   - Medium confidence ({0.5 <= confidence < 0.7}): Carefully evaluate alternatives
   - Low confidence ({confidence < 0.5}): Perform deep analysis of all options

3. Make final decision based on:
   - Primary focus of the procurement
   - Industry best practices
   - Regulatory compliance requirements
   - Category hierarchy appropriateness

RESPONSE FORMAT:
Return only a JSON object with the final category_id and category_name.
Example: {{"category_id": 123, "category_name": "Information Technology"}}"""

    # Add helper function for response parsing
    def _parse_model_response(self, content: str) -> Dict:
        """Parse and validate model response"""
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

            # Ensure proper types
            result["category_name"] = str(result["category_name"])
            result["category_id"] = int(float(result["category_id"]))

            return result
        except Exception as e:
            raise ValueError(f"Error parsing model response: {str(e)}")
