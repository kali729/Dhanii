import requests

def get_branch_address(ifsc_code):
    try:
        # API URL to fetch details by IFSC code
        url = f"https://ifsc.razorpay.com/{ifsc_code}"
        
        # Send a GET request to the API
        response = requests.get(url)
        
        # If the response status code is 200 (OK), parse the response
        if response.status_code == 200:
            data = response.json()
            
            # Fetch and print branch details
            branch_name = data.get("BRANCH")
            bank_name = data.get("BANK")
            address = data.get("ADDRESS")
            
            print(f"Bank: {bank_name}")
            print(f"Branch: {branch_name}")
            print(f"Address: {address}")
        else:
            print("Invalid IFSC code or no information available.")
    
    except Exception as e:
        print(f"Error: {e}")

# Example usage
ifsc_code = input("Enter IFSC Code: ").upper()
get_branch_address(ifsc_code)
