# import streamlit as st
#
# def main():
#     st.markdown(
#         """
#         <style>
#         .reportview-container {
#             background-color: #FFCAD4;
#         }
#
#         .sidebar .sidebar-content {
#             background-color: #0C359E;
#             color: white;
#         }
#
#         .Widget>label {
#             color: white;
#         }
#
#         .stTextArea>div>div>textarea {
#             background-color: #344955;
#             color: white;
#             border-color: #cccccc;
#         }
#
#         .stTextInput>div>div>input {
#             background-color: #344955;
#             color: white;
#             border-color: #cccccc;
#         }
#
#         .stButton>button {
#             background-color: #FFD23F;
#             color: black;
#         }
#
#         .stFileUploader>div>div {
#             background-color: #FFD23F;
#             border-color: #cccccc;
#         }
#         </style>
#         """,
#         unsafe_allow_html=True
#     )
#
#     st.title("My Streamlit App")
#
#     # Sidebar
#     st.sidebar.title("Sidebar")
#     st.sidebar.markdown("This is the sidebar content.")
#
#     # Main content
#     st.write("Main content goes here.")
#
# if __name__ == "__main__":
#     main()
################################################################

import pandas as pd
import uploading


# Provided variable
table_data = sold_table
lines = table_data.strip().split('\n')

# Initialize a list to store the count of elements in each tab
tab_counts = []

# Iterate over each line
for line in lines:
    # Split the line by "|"
    tabs = line.split("|")
    # Iterate over each tab
    for tab in tabs:
        # Count the number of elements in the tab
        elements_count = len(tab.split())
        # Append the count to the tab_counts list
        tab_counts.append(elements_count)

# Find the most common count of elements among all tabs
target_count = max(set(tab_counts), key=tab_counts.count)

# Initialize a list to store potential table tabs
potential_table_tabs = []

# Iterate over each line again to identify potential table tabs
for line in lines:
    # Split the line by "|"
    tabs = line.split("|")
    # Iterate over each tab
    for tab in tabs:
        # Count the number of elements in the tab
        elements_count = len(tab.split())
        # If the count matches the target_count, consider it as a potential table tab
        if elements_count == target_count:
            potential_table_tabs.append(tab)

# Print potential table tabs
for tab in potential_table_tabs:
    print(tab)

