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
from sklearn.metrics.pairwise import cosine_similarity
import torch

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
    def __init__(self, use_test_api=False, embedding_model="paraphrase-MiniLM-L6-v2"):
        """Initialize with configurable embedding model"""
        try:
            # Disable progress bars for transformers
            logging.getLogger("sentence_transformers").setLevel(logging.WARNING)
            logging.getLogger("transformers").setLevel(logging.WARNING)

            # Configure embedding model
            self.embedding_model_name = embedding_model
            self.sentence_model = self._initialize_embedding_model(embedding_model)
            self.nlp = spacy.load("en_core_web_sm")
            
            # Initialize caches
            self.api_cache = TTLCache(maxsize=API_CACHE_SIZE, ttl=API_CACHE_TTL)
            self.embedding_cache = TTLCache(maxsize=EMBEDDING_CACHE_SIZE, ttl=API_CACHE_TTL)
            self.processed_bids = set()
            
            # Domain-specific embeddings
            self.domain_embeddings = {}
            self._load_domain_embeddings()
            
        except Exception as e:
            logger.error(f"Error loading AI models: {str(e)}")
            raise

        # API endpoints based on environment
        if use_test_api:
            self.API_ENDPOINTS = {
                "notice": "http://64.227.157.66/api/getStateNotices",
                "category": "http://64.227.157.66/api/getCategories", 
                "agency": "http://64.227.157.66/api/getStateAgencies",
                "state": "http://64.227.157.66/api/getStates",
            }
        else:
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
        """Process DataFrame with parallel execution and contextual matching"""
        try:
            # Pre-allocate new columns with default values
            df['API Category'] = None
            df['API Category ID'] = None
            df['API Notice Type'] = None
            df['API Agency'] = None
            df['API State'] = None

            # Process chunks in parallel
            chunk_size = max(1, len(df) // (os.cpu_count() or 1))
            chunks = [df[i:i + chunk_size] for i in range(0, len(df), chunk_size)]

            def process_chunk(chunk_df: pd.DataFrame) -> pd.DataFrame:
                """Process a chunk of the DataFrame"""
                try:
                    result_df = pd.DataFrame(index=chunk_df.index)
                    
                    for idx, row in chunk_df.iterrows():
                        try:
                            # Build complete bid context
                            title = str(row.get('Title', row.get('Solicitation Title', ''))).strip()
                            description = str(row.get('Description', '')).strip()
                            agency = str(row.get('Agency', '')).strip()
                            notice_type = str(row.get('Notice Type', '')).strip()
                            url = str(row.get('Bid Detail Page URL', '')).strip()
                            category = str(row.get('Category', '')).strip()

                            # Skip empty rows
                            if not title and not description:
                                continue

                            # Match category
                            category_match = self.find_best_category_match(
                                title, description, category, categories
                            )
                            if category_match and category_match[0]:
                                result_df.at[idx, 'API Category'] = category_match[0]
                                result_df.at[idx, 'API Category ID'] = category_match[1]

                            # Match notice type - try both title and notice_type field
                            combined_text = f"{title} {notice_type} {description}"
                            notice_match = self.determine_notice_type(combined_text, notice_types)
                            if not notice_match or not notice_match[0]:
                                # Fallback: Try just the notice_type field if it exists
                                if notice_type:
                                    notice_match = self.determine_notice_type(notice_type, notice_types)
                            if notice_match and notice_match[0]:
                                result_df.at[idx, 'API Notice Type'] = notice_match[0]
                            else:
                                # Default to "Request For Proposal" if no match found
                                default_type = next(
                                    (nt for nt in notice_types if nt["name"].lower() == "request for proposal"),
                                    next((nt for nt in notice_types), None)
                                )
                                if default_type:
                                    result_df.at[idx, 'API Notice Type'] = default_type["name"]

                            # Match agency - try both agency field and URL
                            agency_match = self.find_best_agency_match(agency, url, agencies)
                            if not agency_match or not agency_match[0]:
                                # Try matching from title if no match found
                                agency_match = self.find_best_agency_match(title, url, agencies)
                            if agency_match and agency_match[0]:
                                result_df.at[idx, 'API Agency'] = agency_match[0]

                        except Exception as e:
                            logger.error(f"Error processing row {idx}: {str(e)}")
                            continue

                    return result_df

                except Exception as e:
                    logger.error(f"Error processing chunk: {str(e)}")
                    return pd.DataFrame()

            # Process chunks in parallel with progress tracking
            print("\nProcessing data in parallel...")
            with ThreadPoolExecutor() as executor:
                futures = [executor.submit(process_chunk, chunk) for chunk in chunks]
                results_dfs = []
                for i, future in enumerate(as_completed(futures)):
                    try:
                        result = future.result()
                        results_dfs.append(result)
                        print(f"\rProgress: {((i+1)/len(futures))*100:.1f}%", end="")
                    except Exception as e:
                        logger.error(f"Error in future {i}: {str(e)}")
            print("\nProcessing complete!")

            # Combine results
            if results_dfs:
                results_df = pd.concat(results_dfs)
                # Update original DataFrame with results
                for col in ['API Category', 'API Category ID', 'API Notice Type', 'API Agency', 'API State']:
                    df[col] = results_df[col]

                # Fill empty notice types with default
                if 'API Notice Type' in df.columns and df['API Notice Type'].isna().any():
                    default_type = next(
                        (nt["name"] for nt in notice_types if nt["name"].lower() == "request for proposal"),
                        next((nt["name"] for nt in notice_types), None)
                    )
                    if default_type:
                        df['API Notice Type'].fillna(default_type, inplace=True)

            return df

        except Exception as e:
            logger.error(f"Error in parallel processing: {str(e)}")
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
        """Make HTTP request with proper error handling and retry logic"""
        max_retries = 3
        retry_delay = 5  # seconds
        
        for attempt in range(max_retries):
            try:
                response = requests.post(
                    url,
                    json=params or {},
                    verify=False,
                    timeout=60,  # Increased timeout to 60 seconds
                    headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                        "Connection": "keep-alive"
                    }
                )
                response.raise_for_status()
                return response
                
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    logger.warning(f"Request timed out. Retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                    sleep(retry_delay)
                    continue
                else:
                    logger.error("Request timed out after all retries")
                    raise
                
            except requests.exceptions.ConnectionError as e:
                if attempt < max_retries - 1:
                    logger.warning(f"Connection error. Retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                    sleep(retry_delay)
                    continue
                else:
                    logger.error("Connection failed after all retries")
                    raise
                
            except Exception as e:
                logger.error(f"Request error: {str(e)}")
                raise
                
        raise requests.exceptions.RequestException("Max retries exceeded")

    def fetch_api_data(self, endpoint_key: str, params: Dict = None) -> List[Dict]:
        """Fetch data from API with proper error handling"""
        try:
            url = self.API_ENDPOINTS[endpoint_key]
            response = self.make_request(url, params)
            data = response.json()

            # Extract data from response
            if isinstance(data, dict) and "data" in data:
                items = data["data"]
            else:
                items = data if isinstance(data, list) else []

            # Transform data based on endpoint
            transformed_data = []
            
            if endpoint_key == "category":
                for item in items:
                    if isinstance(item, dict):
                        transformed_data.append({
                            'name': item.get('category_name', ''),
                            'id': item.get('category_id'),
                            'raw_data': item
                        })
            
            elif endpoint_key == "notice":
                for item in items:
                    if isinstance(item, dict):
                        transformed_data.append({
                            'name': item.get('notice_type', ''),
                            'id': item.get('notice_id'),
                            'raw_data': item
                        })
            
            elif endpoint_key == "agency":
                for item in items:
                    if isinstance(item, dict):
                        transformed_data.append({
                            'name': item.get('agency_name', ''),
                            'id': item.get('agency_id'),
                            'raw_data': item
                        })
            
            elif endpoint_key == "state":
                for item in items:
                    if isinstance(item, dict):
                        transformed_data.append({
                            'name': item.get('state_name', ''),
                            'id': item.get('state_id'),
                            'raw_data': item
                        })

            logger.info(f"Retrieved {len(transformed_data)} items from {endpoint_key} API")
            return transformed_data

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
        """Prepare category embeddings with error handling"""
        try:
            print("Preparing category embeddings...")
            
            # Initialize empty embeddings array
            self.category_embeddings = np.array([])
            
            # Check if we have categories
            if not self.api_categories:
                logger.error("No categories available for embedding preparation")
                return False

            # Extract category texts and prepare embeddings
            category_texts = []
            valid_categories = []
            
            for category in self.api_categories:
                # Get category name from either raw_data or direct name field
                name = (category.get('raw_data', {}).get('category_name') or 
                       category.get('name') or 
                       category.get('category_name', ''))
                
                if name:
                    category_texts.append(str(name))
                    valid_categories.append({
                        'name': name,
                        'id': (category.get('raw_data', {}).get('category_id') or 
                              category.get('id') or 
                              category.get('category_id'))
                    })

            if not category_texts:
                logger.warning("No valid category texts found for embeddings")
                return False

            # Store valid categories
            self.api_categories = valid_categories

            # Generate embeddings for all category texts at once
            self.category_embeddings = self.sentence_model.encode(
                category_texts,
                show_progress_bar=False,
                convert_to_numpy=True
            )
            
            print(f"✅ Prepared embeddings for {len(category_texts)} categories")
            return True

        except Exception as e:
            logger.error(f"Error preparing embeddings: {str(e)}")
            return False

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

    def find_best_category_match(self, title: str, description: str, current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Enhanced category matching with improved preprocessing"""
        try:
            # Check if we have valid embeddings
            if not hasattr(self, 'category_embeddings') or len(self.category_embeddings) == 0:
                logger.warning("No category embeddings available for matching")
                return None, None

            # Clean and normalize inputs using enhanced preprocessing
            title, description = self._preprocess_bid_text(title, description)
            current_category, category_codes = self._preprocess_category_code(current_category)

            if not title and not description:
                return None, None

            # Generate bid identifier for duplicate detection
            bid_id = self._generate_bid_identifier(title, description)
            if bid_id in self.processed_bids:
                logger.info(f"Duplicate bid detected: {title[:50]}...")
                return None, None
            self.processed_bids.add(bid_id)

            # Domain-specific category patterns
            domain_patterns = {
                "software": {
                    "keywords": ["software", "license", "application", "system", "digital", "cloud", "saas", "platform"],
                    "codes": ["20", "208", "209", "43"],
                    "weight": 1.2
                },
                "construction": {
                    "keywords": ["construction", "building", "facility", "infrastructure", "renovation"],
                    "codes": ["90", "91", "236", "237"],
                    "weight": 1.15
                },
                "professional_services": {
                    "keywords": ["consulting", "professional services", "advisory", "management services"],
                    "codes": ["91", "541", "611"],
                    "weight": 1.1
                },
                "equipment": {
                    "keywords": ["equipment", "machinery", "hardware", "apparatus", "device"],
                    "codes": ["23", "333", "334"],
                    "weight": 1.1
                },
                "maintenance": {
                    "keywords": ["maintenance", "repair", "servicing", "support", "preventive"],
                    "codes": ["81", "811", "238"],
                    "weight": 1.05
                }
            }

            # Get embeddings for combined text with weighted components
            combined_text = f"{title} {title} {title} {description} {current_category}"  # Title weighted 3x
            text_embedding = self.get_embedding_cached(combined_text)

            # Calculate base similarities using vectorized operations
            similarities = np.dot(self.category_embeddings, text_embedding) / (
                np.linalg.norm(self.category_embeddings, axis=1) * np.linalg.norm(text_embedding)
            )

            # Apply domain-specific scoring
            for idx, cat in enumerate(self.api_categories):
                cat_name = cat["name"].lower()
                score_multiplier = 1.0

                # Check domain patterns
                for domain, patterns in domain_patterns.items():
                    domain_match = False
                    
                    # Keyword matching
                    keyword_matches = sum(1 for kw in patterns["keywords"] 
                                       if kw in title or kw in description)
                    if keyword_matches > 0:
                        domain_match = True
                        score_multiplier *= (1 + (0.1 * keyword_matches * patterns["weight"]))

                    # Code matching
                    if any(code.startswith(tuple(patterns["codes"])) for code in category_codes):
                        domain_match = True
                        score_multiplier *= patterns["weight"]

                    if domain_match:
                        similarities[idx] *= score_multiplier

                # Direct code matching
                if category_codes:
                    cat_code_match = any(code in cat_name for code in category_codes)
                    if cat_code_match:
                        similarities[idx] *= 1.3
                    elif any(code[:2] in cat_name[:2] for code in category_codes):
                        similarities[idx] *= 1.2

                # Word overlap scoring
                title_words = set(title.split())
                cat_words = set(cat_name.split())
                word_overlap = len(title_words & cat_words)
                if word_overlap > 0:
                    similarities[idx] *= (1 + (0.05 * word_overlap))

            # Get top candidates
            top_k = 5
            top_indices = np.argpartition(similarities, -top_k)[-top_k:]
            candidates = [(idx, similarities[idx]) for idx in top_indices]

            # Additional validation for top candidates
            for idx, score in candidates:
                cat = self.api_categories[idx]
                cat_name = cat["name"].lower()

                # Check for exact matches
                if title.strip() == cat_name or current_category.strip() == cat_name:
                    similarities[idx] *= 1.5

                # Check for substring matches
                elif title.strip() in cat_name or cat_name in title.strip():
                        similarities[idx] *= 1.3

            # Get best match after all adjustments
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            # Dynamic confidence threshold based on match quality
            base_threshold = 0.3
            if category_codes:  # If we have category codes, be more strict
                base_threshold = 0.4
            
            # Adjust threshold based on match quality
            if best_score > 0.8:  # Very strong match
                base_threshold *= 0.8
            elif best_score < 0.5:  # Weak match
                base_threshold *= 1.2

            if best_score < base_threshold:
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

    def determine_notice_type(self, text: str, notice_types: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Enhanced notice type matching with improved pattern recognition"""
        try:
            if not text or not notice_types:
                return None, None

            # Clean and normalize input text
            text = str(text).strip().lower()
            
            # API notice types from response
            API_NOTICE_TYPES = {
                'Request For Proposal': 11,
                'Sources Sought / RFI': 13,
                'Award Notice': 16
            }
            
            # Comprehensive notice type mappings based on common variations
            notice_mappings = {
                # Request For Proposal variations
                'request for proposal': 'Request For Proposal',
                'request for proposals': 'Request For Proposal',
                'rfp': 'Request For Proposal',
                'invitation to bid': 'Request For Proposal',
                'invitation for bid': 'Request For Proposal',
                'ifb': 'Request For Proposal',
                'double envelope proposal': 'Request For Proposal',
                'quick quote': 'Request For Proposal',
                'bid invitation': 'Request For Proposal',
                
                # Sources Sought / RFI variations
                'request for information': 'Sources Sought / RFI',
                'rfi': 'Sources Sought / RFI',
                'information & pricing': 'Sources Sought / RFI',
                'request for information & pricing': 'Sources Sought / RFI',
                'statement of qualification': 'Sources Sought / RFI',
                'soq': 'Sources Sought / RFI',
                'sources sought': 'Sources Sought / RFI',
                
                # Award Notice variations
                'award notice': 'Award Notice',
                'notice of award': 'Award Notice',
                'contract award': 'Award Notice',
                'awarded': 'Award Notice'
            }

            # First try exact mapping from Notice Type field
            for key, value in notice_mappings.items():
                if key in text:
                    return value, API_NOTICE_TYPES.get(value)

            # Default to "Request For Proposal" for common bid indicators
            if any(word in text for word in ['bid', 'solicitation', 'tender', 'proposal', 'quote']):
                return 'Request For Proposal', API_NOTICE_TYPES.get('Request For Proposal')

            return None, None

        except Exception as e:
            logger.error(f"Error determining notice type: {str(e)}")
            return None, None

    def find_best_agency_match(self, agency_name: str, bid_url: str, agencies: List[Dict]) -> Tuple[Optional[str], Optional[str]]:
        """Enhanced agency matching with improved preprocessing and context matching"""
        try:
            if not agency_name and not bid_url:
                return None, None

            # Clean and normalize agency name
            agency_name = str(agency_name).strip().lower() if pd.notna(agency_name) else ""
            bid_url = str(bid_url).strip().lower() if pd.notna(bid_url) else ""

            # Common abbreviations and normalizations
            agency_mappings = {
                'fdot': 'florida department of transportation',
                'dot': 'department of transportation',
                'doh': 'department of health',
                'doe': 'department of education',
                'dor': 'department of revenue',
                'dpw': 'department of public works',
                'isd': 'independent school district',
                'sd': 'school district',
                'sch': 'school',
                'dept': 'department',
                'dist': 'district',
                'auth': 'authority',
                'comm': 'commission',
                'cnty': 'county',
                'co': 'county',
                'univ': 'university',
                'tech': 'technical'
            }

            # Generate agency variations
            agency_variations = set()
            if agency_name:
                # Clean up agency name
                cleaned_name = agency_name.lower()
                cleaned_name = re.sub(r'[^\w\s-]', ' ', cleaned_name)  # Remove special chars except hyphen
                cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()  # Normalize whitespace
                agency_variations.add(cleaned_name)

                # Add variations without common words
                words = cleaned_name.split()
                filtered_words = [w for w in words if len(w) > 2 and w not in {'the', 'and', 'of', 'for', 'to', 'in', 'at'}]
                if filtered_words:
                    agency_variations.add(' '.join(filtered_words))

                # Add mapped variations
                for abbr, full in agency_mappings.items():
                    if abbr in cleaned_name:
                        new_name = cleaned_name.replace(abbr, full)
                        agency_variations.add(new_name)
                    if full in cleaned_name:
                        new_name = cleaned_name.replace(full, abbr)
                        agency_variations.add(new_name)

            # Extract agency from URL if available
            if bid_url:
                url_agency = re.sub(r'https?://(?:www\.)?([^/]+)', r'\1', bid_url)
                url_agency = re.sub(r'\.(?:gov|com|org|edu|net).*$', '', url_agency)
                url_agency = url_agency.replace('.', ' ').strip()
                if url_agency:
                    agency_variations.add(url_agency)

            # Try fuzzy matching with lower threshold
            best_match = None
            best_score = 0
            
            for agency in agencies:
                api_agency_name = agency.get('agency_name', '').lower().strip()
                api_agency_id = str(agency.get('agency_id', ''))
                
                if not api_agency_name:
                    continue

                # Try each variation
                for agency_var in agency_variations:
                    # Calculate similarity scores
                    token_ratio = fuzz.token_set_ratio(agency_var, api_agency_name)
                    partial_ratio = fuzz.partial_ratio(agency_var, api_agency_name)
                    
                    # Combined score with weights
                    score = (token_ratio * 0.7 + partial_ratio * 0.3) / 100.0
                    
                    # Boost score if there are exact word matches
                    agency_words = set(agency_var.split())
                    api_words = set(api_agency_name.split())
                    common_words = agency_words & api_words
                    if common_words:
                        score *= (1 + 0.1 * len(common_words))

                    if score > best_score:
                        best_score = score
                        best_match = (agency.get('agency_name'), api_agency_id)

            # Return match if score is above threshold
            if best_match and best_score > 0.6:  # Lowered threshold from 0.85
                return best_match

            return None, None

        except Exception as e:
            logger.error(f"Error in agency matching: {str(e)}")
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

    def extract_from_description_or_attachment(self, description: str, field_type: str, attachments: str = None) -> str:
        """Extract missing field data from description or attachments."""
        try:
            # Initialize patterns based on field type
            patterns = {
                'agency': [
                    r'(?i)agency:\s*([^,\n]+)',
                    r'(?i)department of\s+([^,\n]+)',
                    r'(?i)([\w\s]+(?:agency|department|office|bureau))',
                ],
                'category': [
                    r'(?i)category:\s*([^,\n]+)',
                    r'(?i)type:\s*([^,\n]+)',
                    r'(?i)procurement of\s+([^,\n]+)',
                    r'(?i)(?:seeking|requesting)\s+([^,\n]+)',
                ],
                'state': [
                    r'(?i)state:\s*([^,\n]+)',
                    r'(?i)location:\s*([^,\n]+)',
                    r'(?i)(?:in|at)\s+(Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|New\s+Hampshire|New\s+Jersey|New\s+Mexico|New\s+York|North\s+Carolina|North\s+Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode\s+Island|South\s+Carolina|South\s+Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|West\s+Virginia|Wisconsin|Wyoming)',
                ]
            }

            # First try description
            if description:
                for pattern in patterns[field_type]:
                    matches = re.findall(pattern, description)
                    if matches:
                        extracted = matches[0].strip()
                        if len(extracted) > 3:  # Avoid very short matches
                            return extracted

            # Then try attachments if provided
            if attachments:
                attachment_list = attachments.split(',')
                for attachment in attachment_list:
                    attachment = attachment.strip()
                    # Look for field type in attachment name
                    if field_type.lower() in attachment.lower():
                        return attachment.split('.')[0].strip()

            return ""

        except Exception as e:
            logger.error(f"Error extracting {field_type} from description: {str(e)}")
            return ""

    def process_row(self, row: pd.Series, processor) -> Dict[str, Any]:
        """Process a single row with fallback to description/attachments."""
        try:
            title = str(row.get('Solicitation Title', row.get('Title', '')))
            description = str(row.get('Description', ''))
            attachments = str(row.get('Attachments', ''))

            # Get or extract agency name
            agency_name = str(row.get('Agency', ''))
            if not agency_name:
                agency_name = self.extract_from_description_or_attachment(description, 'agency', attachments)

            # Get or extract category
            original_category = str(row.get('Category', ''))
            if not original_category:
                original_category = self.extract_from_description_or_attachment(description, 'category', attachments)

            # Get or extract state
            state_name = str(row.get('State', ''))
            if not state_name:
                state_name = self.extract_from_description_or_attachment(description, 'state', attachments)

            # Process matches with extracted data
            return {
                'category': processor.find_best_category_match_ensemble(
                    title, description, original_category, processor.api_categories
                ),
                'notice': processor.determine_notice_type(f"{title} {description}", processor.api_notice_types),
                'agency': processor.find_best_agency_match(agency_name, row.get('Bid Detail Page URL', ''), processor.api_agencies),
                'state': processor.find_state_match(description, agency_name, row.get('Bid Detail Page URL', ''), processor.api_states)
            }

        except Exception as e:
            logger.error(f"Error processing row: {str(e)}")
            return {}

    def find_best_category_match_ensemble(self, title: str, description: str, current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Ensemble method combining multiple matching approaches for better accuracy"""
        try:
            # Clean and normalize inputs
            title = str(title).strip().lower() if pd.notna(title) else ""
            description = str(description).strip().lower() if pd.notna(description) else ""
            current_category = str(current_category).strip().lower() if pd.notna(current_category) else ""

            if not title and not description:
                return None, None

            # Track matches from different methods
            matches = []
            
            # 1. Get match from primary domain-specific method
            primary_match = self.find_best_category_match(title, description, current_category, categories)
            if primary_match[0]:
                matches.append(("primary", primary_match, 0.4))  # 40% weight

            # 2. Get match using semantic similarity only
            semantic_match = self._get_semantic_match(title, description, current_category, categories)
            if semantic_match[0]:
                matches.append(("semantic", semantic_match, 0.3))  # 30% weight

            # 3. Get match using code-based matching
            code_match = self._get_code_match(current_category, categories)
            if code_match[0]:
                matches.append(("code", code_match, 0.2))  # 20% weight

            # 4. Get match using keyword-based matching
            keyword_match = self._get_keyword_match(title, description, categories)
            if keyword_match[0]:
                matches.append(("keyword", keyword_match, 0.1))  # 10% weight

            if not matches:
                return None, None

            # Calculate weighted votes
            category_scores = defaultdict(float)
            for method, (name, id_), weight in matches:
                if name and id_:
                    category_scores[(name, id_)] += weight

            # Get category with highest weighted score
            if category_scores:
                best_match = max(category_scores.items(), key=lambda x: x[1])
                confidence = best_match[1]
                
                # Apply stricter threshold for ensemble
                if confidence >= 0.3:  # At least 30% of weighted votes
                    return best_match[0]

            return None, None

        except Exception as e:
            logger.error(f"Error in ensemble category matching: {str(e)}")
            return None, None

    def _get_semantic_match(self, title: str, description: str, current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Get match based on pure semantic similarity"""
        try:
            # Combine text with weights
            combined_text = f"{title} {title} {description} {current_category}"
            text_embedding = self.get_embedding_cached(combined_text)

            # Calculate similarities
            similarities = np.dot(self.category_embeddings, text_embedding) / (
                np.linalg.norm(self.category_embeddings, axis=1) * np.linalg.norm(text_embedding)
            )

            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            if best_score >= 0.6:  # Higher threshold for pure semantic matching
                return categories[best_idx]["name"], categories[best_idx]["id"]

            return None, None

        except Exception as e:
            logger.error(f"Error in semantic matching: {str(e)}")
            return None, None

    def _get_code_match(self, current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Get match based on category codes"""
        try:
            if not current_category:
                return None, None

            # Extract codes from current category
            codes = re.findall(r'(\d+(?:\.\d+)?)', current_category)
            if not codes:
                return None, None

            # Look for exact code matches first
            for code in codes:
                for cat in categories:
                    cat_name = cat["name"].lower()
                    if code in cat_name:
                        return cat["name"], cat["id"]

            # Try prefix matching
            for code in codes:
                code_prefix = code[:2]
                for cat in categories:
                    cat_name = cat["name"].lower()
                    if cat_name.startswith(code_prefix):
                        return cat["name"], cat["id"]

            return None, None

        except Exception as e:
            logger.error(f"Error in code matching: {str(e)}")
            return None, None

    def _get_keyword_match(self, title: str, description: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Get match based on keyword analysis"""
        try:
            # Extract key phrases from title and description
            title_doc = self.nlp(title)
            desc_doc = self.nlp(description) if description else None

            # Get noun phrases and named entities
            key_phrases = set()
            for doc in [title_doc, desc_doc]:
                if doc:
                    # Add noun phrases
                    key_phrases.update(" ".join(chunk.text.lower() for chunk in doc.noun_chunks))
                    # Add named entities
                    key_phrases.update(ent.text.lower() for ent in doc.ents)

            # Score categories based on keyword matches
            best_score = 0
            best_match = None

            for cat in categories:
                cat_name = cat["name"].lower()
                score = 0

                # Check each key phrase against category name
                for phrase in key_phrases:
                    if phrase in cat_name or cat_name in phrase:
                        score += 1
                    elif fuzz.partial_ratio(phrase, cat_name) > 80:
                        score += 0.5

                if score > best_score:
                    best_score = score
                    best_match = (cat["name"], cat["id"])

            if best_score >= 2:  # Require at least 2 strong matches
                return best_match

            return None, None

        except Exception as e:
            logger.error(f"Error in keyword matching: {str(e)}")
            return None, None

    def find_best_category_match_similarity(self, title: str, description: str, current_category: str, categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Match categories using enhanced embeddings"""
        try:
            # Clean and normalize inputs
            title, description = self._preprocess_bid_text(title, description)
            current_category, _ = self._preprocess_category_code(current_category)

            if not title and not description:
                return None, None

            # Generate bid identifier
            bid_id = self._generate_bid_identifier(title, description)
            if bid_id in self.processed_bids:
                return None, None
            self.processed_bids.add(bid_id)

            # Get enhanced embedding with domain boosting
            combined_text = f"{title} {title} {description} {current_category}"
            text_embedding = self.get_enhanced_embedding(combined_text)

            # Calculate similarities
            similarities = np.dot(self.category_embeddings, text_embedding) / (
                np.linalg.norm(self.category_embeddings, axis=1) * np.linalg.norm(text_embedding)
            )

            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]

            if best_score < 0.3:
                return None, None

            return self.api_categories[best_idx]["name"], self.api_categories[best_idx]["id"]

        except Exception as e:
            logger.error(f"Error in similarity matching: {str(e)}")
            return None, None

    def _preprocess_text(self, text: str, remove_stopwords: bool = True, lemmatize: bool = True) -> str:
        """Enhanced text preprocessing with advanced cleaning and normalization"""
        try:
            if not text or not isinstance(text, str):
                return ""

            # Convert to lowercase and strip
            text = text.lower().strip()

            # Remove URLs
            text = re.sub(r'http[s]?://\S+', '', text)

            # Remove email addresses
            text = re.sub(r'[\w\.-]+@[\w\.-]+', '', text)

            # Remove special characters but keep meaningful punctuation
            text = re.sub(r'[^a-zA-Z0-9\s\.\-\/]', ' ', text)

            # Normalize whitespace
            text = ' '.join(text.split())

            # Process with spaCy for advanced NLP
            doc = self.nlp(text)

            # Initialize processed tokens list
            processed_tokens = []

            for token in doc:
                # Skip if it's a stopword and we want to remove them
                if remove_stopwords and token.is_stop:
                    continue

                # Skip if it's punctuation or whitespace
                if token.is_punct or token.is_space:
                    continue

                # Get the base form if lemmatization is requested
                if lemmatize:
                    token_text = token.lemma_
                else:
                    token_text = token.text

                # Additional cleaning
                token_text = re.sub(r'[^a-z0-9\s\-]', '', token_text)
                
                if token_text:
                    processed_tokens.append(token_text)

            # Join tokens back into text
            processed_text = ' '.join(processed_tokens)

            # Remove redundant whitespace
            processed_text = ' '.join(processed_text.split())

            return processed_text

        except Exception as e:
            logger.error(f"Error in text preprocessing: {str(e)}")
            return text  # Return original text if processing fails

    def _preprocess_category_code(self, category: str) -> Tuple[str, List[str]]:
        """Preprocess category text and extract codes"""
        try:
            if not category:
                return "", []

            # Extract category codes
            codes = re.findall(r'(\d+(?:\.\d+)?)', category)
            
            # Clean category text
            cleaned_category = re.sub(r'\([^)]*\)', '', category)  # Remove parentheses and contents
            cleaned_category = re.sub(r'\s*-\s*', ' - ', cleaned_category)  # Normalize dashes
            cleaned_category = re.sub(r'\s+', ' ', cleaned_category)  # Normalize spaces
            cleaned_category = cleaned_category.strip()

            return cleaned_category, codes

        except Exception as e:
            logger.error(f"Error in category preprocessing: {str(e)}")
            return category, []

    def _preprocess_agency_name(self, agency: str) -> str:
        """Preprocess agency names with specific rules"""
        try:
            if not agency:
                return ""

            # Convert to lowercase and strip
            agency = agency.lower().strip()

            # Remove common prefixes
            prefixes = [
                'department of', 'office of', 'bureau of', 
                'agency for', 'division of', 'board of'
            ]
            for prefix in prefixes:
                if agency.startswith(prefix):
                    agency = agency[len(prefix):].strip()

            # Remove common suffixes
            suffixes = [
                'department', 'office', 'bureau', 'agency',
                'division', 'board', 'administration'
            ]
            for suffix in suffixes:
                if agency.endswith(suffix):
                    agency = agency[:-len(suffix)].strip()

            # Clean special characters
            agency = re.sub(r'[^\w\s\-\.]', '', agency)
            
            # Normalize whitespace
            agency = ' '.join(agency.split())

            return agency

        except Exception as e:
            logger.error(f"Error in agency preprocessing: {str(e)}")
            return agency

    def _preprocess_bid_text(self, title: str, description: str) -> Tuple[str, str]:
        """Preprocess bid title and description with specific rules"""
        try:
            # Process title
            if title:
                # Remove bid/solicitation numbers
                title = re.sub(r'\b(?:bid|solicitation|rfp|rfi|rfq)\s*#?\s*\d+\b', '', title, flags=re.IGNORECASE)
                
                # Remove dates
                title = re.sub(r'\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', '', title)
                
                # Clean and normalize
                title = self._preprocess_text(title, remove_stopwords=False, lemmatize=False)

            # Process description
            if description:
                # Remove common boilerplate text
                description = re.sub(r'for\s+more\s+information\s+.*$', '', description, flags=re.IGNORECASE)
                description = re.sub(r'please\s+contact\s+.*$', '', description, flags=re.IGNORECASE)
                
                # Clean and normalize
                description = self._preprocess_text(description, remove_stopwords=True, lemmatize=True)

            return title or "", description or ""

        except Exception as e:
            logger.error(f"Error in bid text preprocessing: {str(e)}")
            return title or "", description or ""

    def _initialize_embedding_model(self, model_name: str):
        """Initialize the specified embedding model"""
        try:
            if model_name == "roberta-base":
                from transformers import AutoTokenizer, AutoModel
                tokenizer = AutoTokenizer.from_pretrained("roberta-base")
                model = AutoModel.from_pretrained("roberta-base")
                return (tokenizer, model)
            
            elif model_name == "bert-base-uncased":
                from transformers import BertTokenizer, BertModel
                tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")
                model = BertModel.from_pretrained("bert-base-uncased")
                return (tokenizer, model)
            
            else:
                # Default to sentence-transformers
                return SentenceTransformer(model_name)
            
        except Exception as e:
            logger.error(f"Error initializing embedding model {model_name}: {str(e)}")
            # Fallback to default model
            return SentenceTransformer("paraphrase-MiniLM-L6-v2")

    def _get_embedding(self, text: str) -> np.ndarray:
        """Get embeddings using the configured model"""
        try:
            if isinstance(self.sentence_model, tuple):  # BERT/RoBERTa
                tokenizer, model = self.sentence_model
                inputs = tokenizer(text, return_tensors="pt", padding=True, truncation=True)
                outputs = model(**inputs)
                # Use CLS token embedding
                embedding = outputs.last_hidden_state[:, 0, :].detach().numpy()
                return embedding.squeeze()
            
            else:  # sentence-transformers
                return self.sentence_model.encode([text])[0]
            
        except Exception as e:
            logger.error(f"Error getting embedding: {str(e)}")
            return np.zeros(384)  # Return zero vector as fallback

    def get_enhanced_embedding(self, text: str, domain_boost: bool = True) -> np.ndarray:
        """Get enhanced embeddings with domain-specific boosting"""
        try:
            # Get base embedding
            base_embedding = self._get_embedding(text)
            
            if not domain_boost:
                return base_embedding
            
            # Calculate domain similarities
            domain_scores = {}
            for domain, domain_embedding in self.domain_embeddings.items():
                similarity = np.dot(base_embedding, domain_embedding) / (
                    np.linalg.norm(base_embedding) * np.linalg.norm(domain_embedding)
                )
                domain_scores[domain] = similarity
            
            # Find dominant domain
            dominant_domain = max(domain_scores.items(), key=lambda x: x[1])
            
            if dominant_domain[1] > 0.3:  # Apply domain boost only if strong match
                # Boost embedding in direction of domain
                boost_factor = 0.2  # 20% boost
                domain_embedding = self.domain_embeddings[dominant_domain[0]]
                boosted_embedding = base_embedding + (boost_factor * domain_embedding)
                
                # Normalize
                return boosted_embedding / np.linalg.norm(boosted_embedding)
            
            return base_embedding
            
        except Exception as e:
            logger.error(f"Error getting enhanced embedding: {str(e)}")
            return self._get_embedding(text)

    def _load_domain_embeddings(self):
        """Load domain-specific embeddings"""
        try:
            # Define domain-specific vocabulary
            domain_vocab = {
                "procurement": [
                    "bid", "tender", "solicitation", "proposal", "contract",
                    "vendor", "supplier", "procurement", "acquisition"
                ],
                "construction": [
                    "construction", "building", "renovation", "infrastructure",
                    "contractor", "facility", "installation"
                ],
                "technology": [
                    "software", "hardware", "system", "network", "database",
                    "application", "platform", "cloud", "digital"
                ],
                "services": [
                    "consulting", "maintenance", "support", "professional",
                    "management", "operation", "service"
                ]
            }
            
            # Generate domain embeddings
            for domain, terms in domain_vocab.items():
                domain_text = " ".join(terms)
                self.domain_embeddings[domain] = self._get_embedding(domain_text)
            
        except Exception as e:
            logger.error(f"Error loading domain embeddings: {str(e)}")

    def _get_contextual_embedding(self, bid_data: Dict[str, str], context_window: int = 512) -> np.ndarray:
        """Get contextual embedding considering the entire bid document structure"""
        try:
            # Extract all available bid information
            title = bid_data.get('title', '').strip()
            description = bid_data.get('description', '').strip()
            category = bid_data.get('category', '').strip()
            agency = bid_data.get('agency', '').strip()
            notice_type = bid_data.get('notice_type', '').strip()
            
            # Structure the document with section markers for better context
            doc_sections = [
                f"TITLE: {title}" if title else "",
                f"CATEGORY: {category}" if category else "",
                f"NOTICE TYPE: {notice_type}" if notice_type else "",
                f"AGENCY: {agency}" if agency else "",
                f"DESCRIPTION: {description}" if description else ""
            ]
            
            # Join sections with special markers
            structured_text = " [SEP] ".join(filter(None, doc_sections))
            
            # Truncate to context window while preserving important parts
            if len(structured_text.split()) > context_window:
                # Keep full title and category, truncate description
                important_parts = doc_sections[:2]  # Title and category
                remaining_tokens = context_window - sum(len(part.split()) for part in important_parts)
                
                if description:
                    # Truncate description to fit remaining space
                    words = description.split()
                    if len(words) > remaining_tokens:
                        # Keep start and end of description
                        start_tokens = remaining_tokens // 2
                        end_tokens = remaining_tokens - start_tokens
                        truncated_desc = " ".join(words[:start_tokens] + ["..."] + words[-end_tokens:])
                        doc_sections[-1] = f"DESCRIPTION: {truncated_desc}"
                
                structured_text = " [SEP] ".join(filter(None, doc_sections))
            
            # Get embedding with context-aware processing
            return self._get_contextual_transformer_embedding(structured_text)
            
        except Exception as e:
            logger.error(f"Error in contextual embedding: {str(e)}")
            return self._get_embedding(str(bid_data.get('title', '')))

    def _get_contextual_transformer_embedding(self, text: str) -> np.ndarray:
        """Get embeddings using context-aware transformer processing"""
        try:
            if isinstance(self.sentence_model, tuple):  # BERT/RoBERTa
                tokenizer, model = self.sentence_model
                
                # Tokenize with attention to special tokens
                inputs = tokenizer(
                    text,
                    return_tensors="pt",
                    padding=True,
                    truncation=True,
                    max_length=512,
                    add_special_tokens=True
                )
                
                # Get model outputs with attention weights
                with torch.no_grad():
                    outputs = model(**inputs)
                
                # Use attention-weighted average of token embeddings
                attention_weights = outputs.attentions[-1].mean(dim=1)  # Use last layer
                token_embeddings = outputs.last_hidden_state
                
                # Weight token embeddings by attention
                weighted_embeddings = (token_embeddings * attention_weights.unsqueeze(-1)).sum(dim=1)
                return weighted_embeddings.numpy().squeeze()
                
            else:  # sentence-transformers
                # Use mean pooling for sentence transformers
                return self.sentence_model.encode(
                    text,
                    show_progress_bar=False,
                    convert_to_numpy=True,
                    normalize_embeddings=True
                )
                
        except Exception as e:
            logger.error(f"Error in transformer embedding: {str(e)}")
            return self._get_embedding(text)

    def find_best_category_match_contextual(self, bid_data: Dict[str, str], categories: List[Dict]) -> Tuple[Optional[str], Optional[int]]:
        """Enhanced category matching using contextual embeddings"""
        try:
            # Get contextual embedding for the bid
            bid_embedding = self._get_contextual_embedding(bid_data)
            
            # Get contextual embeddings for categories
            category_embeddings = []
            for category in categories:
                cat_data = {
                    'title': category['name'],
                    'description': category.get('description', ''),
                    'category': category.get('parent_name', '')
                }
                category_embeddings.append(
                    self._get_contextual_embedding(cat_data)
                )
            
            category_embeddings = np.array(category_embeddings)
            
            # Calculate contextual similarities
            similarities = np.dot(category_embeddings, bid_embedding) / (
                np.linalg.norm(category_embeddings, axis=1) * np.linalg.norm(bid_embedding)
            )
            
            # Apply hierarchical boost
            for idx, category in enumerate(categories):
                if category.get('parent_id'):
                    # Boost child categories based on parent match
                    parent_idx = next(
                        (i for i, cat in enumerate(categories) 
                         if cat['id'] == category['parent_id']),
                        None
                    )
                    if parent_idx is not None:
                        parent_similarity = similarities[parent_idx]
                        similarities[idx] *= (1 + 0.2 * parent_similarity)
            
            # Get best match
            best_idx = np.argmax(similarities)
            best_score = similarities[best_idx]
            
            # Apply confidence threshold
            if best_score < 0.4:  # Higher threshold for contextual matching
                return None, None
                
            return categories[best_idx]["name"], categories[best_idx]["id"]
            
        except Exception as e:
            logger.error(f"Error in contextual category matching: {str(e)}")
            return None, None

def process_excel_from_cli(base_path: str = None, use_test_api: bool = False, matching_method: str = "similarity") -> bool:
    """Process Excel files from yesterday's COMPLETED folders"""
    try:
        # Get yesterday's date folder
        from datetime import datetime, timedelta
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        if not base_path:
            base_path = os.path.join(os.getcwd(), yesterday)
        
        if not os.path.exists(base_path):
            print(f"❌ Yesterday's folder not found: {base_path}")
            return False
            
        print(f"\nProcessing Excel files from: {base_path}")
        print(f"Using matching method: {matching_method}")
        
        # Initialize processor with test API flag
        processor = ExcelProcessor(use_test_api=use_test_api)
        
        # Set the matching method
        processor.matching_method = matching_method
        
        # Fetch API data with retries
        print("\n📥 Fetching API data...")
        max_retries = 3
        for attempt in range(max_retries):
            processor.api_categories = processor.fetch_api_data("category")
            processor.api_notice_types = processor.fetch_api_data("notice")
            processor.api_agencies = processor.fetch_api_data("agency")
            processor.api_states = processor.fetch_api_data("state", {"country_id": 10})
            
            if all([processor.api_categories, processor.api_notice_types, 
                   processor.api_agencies, processor.api_states]):
                print("✅ API data fetched successfully")
                break
            elif attempt < max_retries - 1:
                print(f"⚠️ Retry {attempt + 1}/{max_retries} fetching API data...")
                sleep(2)  # Wait before retry
            else:
                print("❌ Failed to fetch API data after multiple attempts")
                return False

        # Prepare embeddings only if we have categories
        if processor.api_categories:
            print("📊 Preparing category embeddings...")
            processor._prepare_embeddings()
            if processor.category_embeddings is not None:
                print("✅ Category embeddings prepared")
            else:
                print("❌ Failed to prepare category embeddings")
                return False
        else:
            print("❌ No categories available for processing")
            return False

        # Find all COMPLETED folders
        success = True
        excel_files_processed = 0
        
        for root, dirs, files in os.walk(base_path):
            if root.endswith('COMPLETED'):
                print(f"\n Processing folder: {root}")
                excel_files = [f for f in files if f.endswith('.xlsx')]
                
                if not excel_files:
                    print("  No Excel files found in this COMPLETED folder")
                    continue
                    
                for excel_file in excel_files:
                    excel_path = os.path.join(root, excel_file)
                    print(f"\n Processing: {excel_file}")
                    
                    try:
                        # Load Excel file
                        df = pd.read_excel(excel_path)
                        total_rows = len(df)
                        print(f"📊 Total rows to process: {total_rows}")

                        # Handle different title column names
                        title_column = 'Solicitation Title' if 'Solicitation Title' in df.columns else 'Title'
                        if title_column not in df.columns:
                            print("❌ No Title or Solicitation Title column found")
                            continue

                        # Get column positions for inserting API columns
                        category_pos = df.columns.get_loc('Category') + 1 if 'Category' in df.columns else len(df.columns)
                        notice_pos = df.columns.get_loc('Notice Type') + 1 if 'Notice Type' in df.columns else len(df.columns)
                        agency_pos = df.columns.get_loc('Agency') + 1 if 'Agency' in df.columns else len(df.columns)
                        state_pos = df.columns.get_loc('State') + 1 if 'State' in df.columns else len(df.columns)

                        # Add API columns
                        if 'API_Category' not in df.columns:
                            df.insert(category_pos, 'API_Category', None)
                            df.insert(category_pos + 1, 'API_Category_ID', None)
                        if 'API_Notice_Type' not in df.columns:
                            df.insert(notice_pos, 'API_Notice_Type', None)
                        if 'API_Agency' not in df.columns:
                            df.insert(agency_pos, 'API_Agency', None)
                        if 'API_State' not in df.columns:
                            df.insert(state_pos, 'API_State', None)

                        # Process each row
                        for index, row in df.iterrows():
                            try:
                                # Calculate progress
                                progress = int(((index + 1) / total_rows) * 100)
                                print(f"\rProcessing row {index + 1}/{total_rows} ({progress}%)", end='')

                                # Get fields for matching
                                title = str(row.get(title_column, ''))
                                description = str(row.get('Description', ''))
                                original_category = str(row.get('Category', ''))
                                agency_name = str(row.get('Agency', ''))
                                bid_url = str(row.get('Bid Detail Page URL', ''))

                                # Match category using selected method
                                if processor.matching_method == "ensemble":
                                    category_match = processor.find_best_category_match_ensemble(
                                        title, description, original_category, processor.api_categories
                                    )
                                elif processor.matching_method == "hybrid":
                                    category_match = processor.find_best_category_match(
                                        title, description, original_category, processor.api_categories
                                    )
                                else:  # Default to similarity method
                                    category_match = processor.find_best_category_match_similarity(
                                        title, description, original_category, processor.api_categories
                                    )

                                if category_match and category_match[0]:
                                    df.at[index, 'API_Category'] = category_match[0]
                                    df.at[index, 'API_Category_ID'] = category_match[1]

                                # Match notice type
                                notice_type = processor.determine_notice_type(
                                    f"{title} {description}", processor.api_notice_types
                                )
                                if notice_type and notice_type[0]:
                                    df.at[index, 'API_Notice_Type'] = notice_type[0]

                                # Match agency
                                agency_match = processor.find_best_agency_match(
                                    agency_name, bid_url, processor.api_agencies
                                )
                                if agency_match and agency_match[0]:
                                    df.at[index, 'API_Agency'] = agency_match[0]

                                # Match state
                                state_match = processor.find_state_match(
                                    description, agency_name, bid_url, processor.api_states
                                )
                                if state_match and state_match[0]:
                                    df.at[index, 'API_State'] = state_match[0]

                            except Exception as e:
                                print(f"\n❌ Error processing row {index + 1}: {str(e)}")
                                continue

                        # Save processed file
                        df.to_excel(excel_path, index=False)
                        print(f"\n✅ Saved processed file: {excel_file}")
                        excel_files_processed += 1

                    except Exception as e:
                        print(f"\n❌ Error processing {excel_file}: {str(e)}")
                        success = False
                        continue

        print(f"\n🎉 Processing complete! Processed {excel_files_processed} Excel files")
        return success

    except Exception as e:
        print(f"\n❌ Error during processing: {str(e)}")
        return False

if __name__ == "__main__":
    import sys
    import argparse
    
    parser = argparse.ArgumentParser(description='Process Excel files with category matching')
    parser.add_argument('--base-path', help='Base path to process files from')
    parser.add_argument('--test-api', action='store_true', help='Use test API endpoints')
    parser.add_argument('--method', choices=['similarity', 'ensemble', 'hybrid'], 
                       default='similarity', help='Category matching method to use')
    parser.add_argument('--embedding-model', 
                       choices=['paraphrase-MiniLM-L6-v2', 'roberta-base', 
                               'bert-base-uncased'],
                       default='paraphrase-MiniLM-L6-v2',
                       help='Embedding model to use')
    
    args = parser.parse_args()
    
    processor = ExcelProcessor(use_test_api=args.test_api, 
                             embedding_model=args.embedding_model)
    success = process_excel_from_cli(args.base_path, processor)
    sys.exit(0 if success else 1)
