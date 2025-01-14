import boto3

s3_client = boto3.client(
    "s3",
    aws_access_key_id="minioadmin",
    aws_secret_access_key="minioadmin",
    endpoint_url="http://localhost:9000",
)

# List buckets to test connection
response = s3_client.list_buckets()
print("Existing buckets:")
for bucket in response["Buckets"]:
    print(f'  {bucket["Name"]}')
