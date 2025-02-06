import pytest
import requests
import json
from pprint import pprint
import urllib3
import warnings
import time
import argparse
from typing import List, Dict

# Suppress SSL verification warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# API endpoints
API_ENDPOINTS = {
    "notice": "https://bidsportal.com/api/getNotice",
    "category": "https://bidsportal.com/api/getCategory",
    "agency": "https://bidsportal.com/api/getAgency",
    "state": "https://bidsportal.com/api/getState",
}

def print_column_values(data: List[Dict], column: str):
    """Print values for a specific column from the API response"""
    if not data:
        print(f"No data found for {column}")
        return

    print(f"\n=== {column.upper()} Values ===\n")
    
    # Map column names to their corresponding data keys
    column_mappings = {
        'agency': 'agency_name',
        'notice': 'notice_type',
        'category': 'category_name',
        'state': 'state_name',
        'agency_id': 'agency_id',
        'notice_id': 'notice_id',
        'category_id': 'category_id',
        'state_id': 'state_id'
    }

    key = column_mappings.get(column.lower())
    if not key:
        print(f"Unknown column: {column}")
        return

    # Print values with their IDs
    for item in data:
        if isinstance(item, dict):
            name = item.get(key, '')
            id_field = key.replace('name', 'id') if 'name' in key else key
            id_value = item.get(id_field, '')
            
            if name:
                if '_id' in column:
                    print(f"{id_value}")
                else:
                    print(f"{name} (ID: {id_value})")

def make_request(url: str, params: dict = None) -> requests.Response:
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
        return response
    except requests.exceptions.SSLError as e:
        print(f"SSL Error: {str(e)}")
        print("Attempting request without SSL verification...")
        return requests.post(url, json=params or {}, verify=False)
    except requests.exceptions.ConnectionError as e:
        print(f"Connection Error: {str(e)}")
        raise
    except requests.exceptions.Timeout as e:
        print(f"Timeout Error: {str(e)}")
        raise
    except requests.exceptions.RequestException as e:
        print(f"Request Error: {str(e)}")
        raise

def get_api_data(endpoint_name: str) -> List[Dict]:
    """Get data from specified API endpoint"""
    url = API_ENDPOINTS.get(endpoint_name)
    if not url:
        print(f"Invalid endpoint: {endpoint_name}")
        return []

    try:
        params = {"country_id": 10} if endpoint_name == "state" else {}
        response = make_request(url, params)
        
        if response.status_code == 200:
            data = response.json()
            if isinstance(data, dict) and "data" in data:
                # Print sample data for debugging
                if data["data"]:
                    print("\nSample API Response Structure:")
                    pprint(data["data"][0])
                return data["data"]
            
        print(f"Error getting data from {endpoint_name} API: {response.status_code}")
        return []

    except Exception as e:
        print(f"Error accessing {endpoint_name} API: {str(e)}")
        return []

def test_api_endpoints(endpoint_filter: str = None, show_column: str = None):
    """Test each API endpoint and optionally show specific column values"""
    if endpoint_filter and show_column:
        # Get and display only the requested column for the specified endpoint
        data = get_api_data(endpoint_filter)
        print_column_values(data, show_column)
        return

    print("\n=== Testing API Endpoints ===\n")

    # Filter endpoints if specified
    endpoints_to_test = {
        k: v
        for k, v in API_ENDPOINTS.items()
        if endpoint_filter is None or k == endpoint_filter
    }

    if endpoint_filter and not endpoints_to_test:
        print(f"❌ Invalid endpoint filter: {endpoint_filter}")
        print(f"Valid endpoints are: {', '.join(API_ENDPOINTS.keys())}")
        return

    for endpoint_name, url in endpoints_to_test.items():
        print(f"\nTesting {endpoint_name.upper()} API:")
        print(f"URL: {url}")

        try:
            # Make the API call with SSL verification disabled
            params = {"country_id": 10} if endpoint_name == "state" else {}
            response = make_request(url, params)

            # Check if the response was successful
            if response.status_code == 200:
                try:
                    # Parse and print the response
                    data = response.json()

                    print("\nResponse Status:", response.status_code)
                    print("\nResponse Headers:")
                    pprint(dict(response.headers))

                    print("\nResponse Data:")
                    pprint(data)

                    # Basic validation of response data
                    if data:  # Check if data is not empty
                        if isinstance(data, dict) and "data" in data:
                            items = data["data"]
                            if items:
                                print(f"\nNumber of items: {len(items)}")
                                print(f"\nSample {endpoint_name} item structure:")
                                pprint(items[0])
                            else:
                                print(
                                    f"\n⚠️ Warning: Empty data array from {endpoint_name} API"
                                )
                        else:
                            print(
                                f"\n⚠️ Warning: Missing 'data' key in {endpoint_name} response"
                            )
                    else:
                        print(f"\n⚠️ Warning: Empty response from {endpoint_name} API")

                except json.JSONDecodeError as e:
                    print(
                        f"\n⚠️ Warning: Invalid JSON response from {endpoint_name} API"
                    )
                    print(
                        f"Response text: {response.text[:200]}..."
                    )  # Print first 200 chars
            else:
                print(
                    f"\n⚠️ Warning: Unexpected status code {response.status_code} from {endpoint_name} API"
                )
                print(
                    f"Response text: {response.text[:200]}..."
                )  # Print first 200 chars

        except Exception as e:
            print(f"\n❌ Error testing {endpoint_name} API: {str(e)}")
            print(f"Type of error: {type(e).__name__}")

        finally:
            print("-" * 50)


def test_api_response_times(endpoint_filter: str = None):
    """Test and display API response times"""

    print("\n=== Testing API Response Times ===\n")

    # Filter endpoints if specified
    endpoints_to_test = {
        k: v
        for k, v in API_ENDPOINTS.items()
        if endpoint_filter is None or k == endpoint_filter
    }

    for endpoint_name, url in endpoints_to_test.items():
        try:
            print(f"\nTesting {endpoint_name.upper()} API response time:")

            # Make the API call and measure response time
            params = {"country_id": 10} if endpoint_name == "state" else {}
            start_time = time.time()
            response = make_request(url, params)
            elapsed_time = time.time() - start_time

            print(f"Response Time: {elapsed_time:.3f} seconds")
            print(f"Status Code: {response.status_code}")

            if elapsed_time > 5:
                print(f"⚠️  Warning: High response time for {endpoint_name} API")

        except Exception as e:
            print(f"❌ Error measuring response time for {endpoint_name}: {str(e)}")

        finally:
            print("-" * 50)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test API endpoints")
    parser.add_argument(
        "--endpoint",
        type=str,
        choices=["notice", "category", "agency", "state"],
        help="Specify a single endpoint to test (notice, category, agency, or state)",
    )
    parser.add_argument(
        "--show-column",
        type=str,
        choices=["agency", "notice", "category", "state", 
                "agency_id", "notice_id", "category_id", "state_id"],
        help="Show only values for the specified column"
    )
    parser.add_argument(
        "--response-times",
        action="store_true",
        help="Test API response times"
    )

    args = parser.parse_args()

    print("\n🚀 Starting API Tests...")

    try:
        test_api_endpoints(args.endpoint, args.show_column)
        if args.response_times:
            test_api_response_times(args.endpoint)
        print("\n✅ All API tests completed!")

    except Exception as e:
        print(f"\n❌ Tests failed: {str(e)}")

    finally:
        print("\n=== Test Run Complete ===")
