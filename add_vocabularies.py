import pandas as pd
import requests

# Load the Excel file
file_path = "OpenCTI vocabularies_example.xlsx"
xls = pd.ExcelFile(file_path)

# GraphQL API endpoint and authentication token
API_URL = "https://somthing.octi.filigran.io/graphql"
HEADERS = {
    "Authorization": "Bearer API_KEY",
    "Content-Type": "application/json"
}

# GraphQL queries
query_existing_vocabularies = """
query GetVocabularies($category: VocabularyCategory!) {
  vocabularies(category: $category) {
    edges {
      node {
        id
        name
        description
      }
    }
  }
}
"""

mutation_delete_query = """
mutation VocabularyDeletionMutation($id: ID!) {
  vocabularyDelete(id: $id)
}
"""

mutation_add_query = """
mutation VocabularyCreationMutation($input: VocabularyAddInput!) {
  vocabularyAdd(input: $input) {
    id
    name
  }
}
"""

# Process each sheet (category) in the Excel file
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    df = df.iloc[1:].reset_index(drop=True)  # Ignore the first row (header) and reset index

    # Fetch existing vocabularies for the category
    response = requests.post(API_URL, json={"query": query_existing_vocabularies, "variables": {"category": sheet_name}}, headers=HEADERS)
    if response.status_code == 200:
        existing_vocabularies = response.json().get("data", {}).get("vocabularies", {}).get("edges", [])

        for vocab in existing_vocabularies:
            vocab_id = vocab["node"]["id"]
            delete_payload = {"query": mutation_delete_query, "variables": {"id": vocab_id}}
            delete_response = requests.post(API_URL, json=delete_payload, headers=HEADERS)
            if delete_response.status_code == 200:
                print(f"Deleted vocabulary: {vocab['node']['name']} from category {sheet_name}")
            else:
                print(f"Failed to delete {vocab['node']['name']}: {delete_response.text}")
    else:
        print(f"Failed to fetch existing vocabularies for category {sheet_name}: {response.text}")
        continue

    # Add new vocabularies
    for index, row in df.iterrows():
        vocab_name = row.iloc[0]  # First column is the vocabulary name
        description = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else ""  # Column B contains descriptions, default to "" if empty

        # Construct input payload
        payload = {
            "query": mutation_add_query,
            "variables": {
                "input": {
                    "name": vocab_name,
                    "description": description,
                    "aliases": [],
                    "category": sheet_name
                }
            }
        }

        # Send request to GraphQL API
        response = requests.post(API_URL, json=payload, headers=HEADERS)

        # Check response
        if response.status_code == 200:
            print(f"Added vocabulary: {vocab_name} under category {sheet_name}")
        else:
            print(f"Failed to add {vocab_name}: {response.text}")
