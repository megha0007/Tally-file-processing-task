import xml.etree.ElementTree as ET
import pandas as pd
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.core.files.storage import default_storage
import os
from datetime import datetime


@csrf_exempt
def process_tally(request):
    if request.method != 'POST':
        return JsonResponse({"error": "Only POST requests are allowed"}, status=405)

    if 'file' not in request.FILES:
        return JsonResponse({"error": "No file uploaded"}, status=400)

    # Save uploaded file temporarily
    xml_file = request.FILES['file']
    temp_file = default_storage.save("temp.xml", xml_file)

    try:
        # Parse the XML and extract transactions
        transactions = parse_tally_xml(temp_file)

        if not transactions:
            return JsonResponse({"message": "No 'Receipt' vouchers found."}, status=200)

        # Generate Excel
        output_file = "response.xlsx"
        generate_excel(transactions, output_file)

        # Serve the file
        with open(output_file, "rb") as f:
            response = HttpResponse(
                f.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            response['Content-Disposition'] = f'attachment; filename="{output_file}"'
            return response

    finally:
        # Clean up
        if os.path.exists(temp_file):
            os.remove(temp_file)
        if os.path.exists(output_file):
            os.remove(output_file)

def parse_tally_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    transactions = []
    child_transactions = {}

    for voucher in root.findall(".//VOUCHER"):
        voucher_number = voucher.find("VOUCHERNUMBER").text if voucher.find("VOUCHERNUMBER") is not None else "NA"
        voucher_type = voucher.find("VOUCHERTYPENAME").text if voucher.find("VOUCHERTYPENAME") is not None else "NA"
        voucher_date = voucher.find("DATE").text if voucher.find("DATE") is not None else "NA"
        ledger_name = voucher.find(".//LEDGERNAME").text if voucher.find(".//LEDGERNAME") is not None else "NA"
        if voucher_date != "NA":
            try:
                # Parse the date in the original format (assuming it's in YYYYMMDD or a different format)
                formatted_date = datetime.strptime(voucher_date, "%Y%m%d").strftime("%d-%m-%Y")
            except ValueError:
                formatted_date = voucher_date  # If there's an error in parsing, keep it as is
        else:
            formatted_date = "NA"

        # Initialize base transaction structure
        transaction = {
            "Date": formatted_date,
            "Voucher Type": voucher_type,
            "Transaction Type": "NA",
            "Vch No.": voucher_number,
            "Ref No.": "NA",
            "Ref Type": "NA",
            "Ref Date": "NA",
            "Debtor": ledger_name,
            "Ref Amount": "0",
            "Amount": voucher.find(".//AMOUNT").text if voucher.find(".//AMOUNT") is not None else "0",
            "Particulars": ledger_name,
            "Vch Type": voucher_type,
            "Amount Verified": "NA",
        }
        # Handle "Parent" transaction type
        if voucher_type == "Receipt":
            # Create the "Parent" transaction
            parent_transaction = transaction.copy()
            parent_transaction["Transaction Type"] = "Parent"
            transactions.append(parent_transaction)
            # Add "Other" transaction type for the same Vch No.
            other_transaction = transaction.copy()
            other_transaction["Transaction Type"] = "Other"
            other_transaction["Amount"] = "-" + transaction["Amount"]
            other_transaction["Ref Amount"] = "NA"
            transactions.append(other_transaction)

            # Check for child transactions (sub-transactions)
            bill_allocations = voucher.findall(".//BILLALLOCATIONS.LIST")
            if bill_allocations:
                # Track added Ref Nos to avoid duplicates
                added_ref_nos = set()

                for bill_allocation in bill_allocations:
                    ref_no = bill_allocation.find("NAME").text if bill_allocation.find("NAME") is not None else "NA"
                    
                    # Skip if this ref_no has already been added (avoiding duplicates)
                    if ref_no != "NA":
                    # Skip if this ref_no has already been added (avoiding duplicates)
                        if ref_no in added_ref_nos:
                            continue

                        # Otherwise, create the child transaction
                        child_transaction = transaction.copy()
                        child_transaction["Transaction Type"] = "Child"
                        child_transaction["Debtor"] = ledger_name
                        child_transaction["Amount"] = "NA"
                        child_transaction["Ref No."] = ref_no
                        child_transaction["Ref Type"] = bill_allocation.find("BILLTYPE").text if bill_allocation.find("BILLTYPE") is not None else "NA"
                        child_transaction["Ref Date"] = bill_allocation.find("DATE").text if bill_allocation.find("DATE") is not None else ""
                        child_transaction["Ref Amount"] = bill_allocation.find("AMOUNT").text if bill_allocation.find("AMOUNT") is not None else "0"
                        # Add the ref_no to the set to ensure it's not added again
                        added_ref_nos.add(ref_no)
                        # Append the child transaction
                        transactions.append(child_transaction)


    # Update "Amount Verified" for Parent transactions
    for transaction in transactions:
      
        if transaction["Transaction Type"] == "Parent":
         
            parent_vch_no = transaction["Vch No."]
            child_ref_amount_sum = sum(
                float(child["Ref Amount"]) for child in transactions if child["Transaction Type"] == "Child" and child["Vch No."] == parent_vch_no
            )
            amount_verified = "Yes" if child_ref_amount_sum == float(transaction["Amount"]) else "No"
            transaction["Amount Verified"] = amount_verified
            transaction["Ref No."] = "NA"
            transaction["Ref Type"] = "NA"
            transaction["Ref Date"] =  "NA"
            transaction["Ref Amount"] = "NA"
             
    return transactions

def generate_excel(transactions, output_file):
    """
    Generate an Excel file from the transactions and overwrite if the file exists.
    """
    # Replace "NA" with blank where necessary
    for transaction in transactions:
        for key, value in transaction.items():
            if value == "NA":
                transaction[key] = "NA"  # Empty cells for NA values

    df = pd.DataFrame(transactions)
    df.to_excel(output_file, index=False, engine='openpyxl')  # Overwrites existing file
