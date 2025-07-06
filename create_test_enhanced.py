#!/usr/bin/env python3

import pandas as pd
from datetime import datetime, timedelta
import random

# Test data with Reviewer column
reviewers = ["John Doe", "Jane Smith", "Bob Johnson", "Alice Chen", "Mike Wilson"]
departments = ["IT", "Finance", "HR", "Sales", "Admin"]
statuses = ["Pending", "In Progress", "Completed"]
applications = ["SAP", "Salesforce", "Office365", "Slack", "Zoom"]

# Create test data
data = []
for i in range(50):
    data.append({
        "User_ID": f"USR-{i+1:04d}",
        "Name": f"Employee{random.randint(1, 30)}",
        "Department": random.choice(departments),
        "Reviewer": random.choice(reviewers),
        "Application": random.choice(applications),
        "Access_Level": random.choice(["Read", "Write", "Admin"]),
        "Request_Date": datetime.now() - timedelta(days=random.randint(0, 30)),
        "Status": random.choice(statuses),
        "Comments": f"Access request {i+1}"
    })

# Create DataFrame and save
df = pd.DataFrame(data)
df.to_excel("user_listing.xlsx", index=False, engine='openpyxl')
print("Test file 'user_listing.xlsx' created!")

# Create sample Word document content
word_content = """Sample Word Document
This would be a .docx file in real scenario.
Created for testing purposes.
"""

# Create sample PDF content  
pdf_content = """Sample Permission PDF
This would be a .pdf file in real scenario.
Contains permission forms for the application.
"""

# Save as text files to simulate documents
with open("TestApp_UserGuide.txt", "w") as f:
    f.write("TestApp User Guide\n" + word_content)

with open("TestApp_permission_form.txt", "w") as f:
    f.write("TestApp Permission Form\n" + pdf_content)

print("Sample document files created (as .txt for testing)")
print("\nNote: In production, these would be:")
print("- TestApp_UserGuide.docx")
print("- TestApp_permission_form.pdf")