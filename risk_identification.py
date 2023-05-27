# -*- coding: utf-8 -*-
"""
Created on Sun May 28 00:29:41 2023

@author: riadh
"""

import pandas as pd
from openpyxl import Workbook
import matplotlib.pyplot as plt

# Define the list of potential cyber security risks
risks = ['Data breaches', 'Hacking attempts', 'Malware or phishing attacks']

# Create an empty DataFrame to store the risk data
risk_data = pd.DataFrame(columns=['Risk', 'Likelihood', 'Impact', 'Mitigation'])

# Loop through each risk and prompt the user to input the likelihood, impact, and recommended mitigation strategy
for risk in risks:
    likelihood = input(f"What is the likelihood of a {risk} occurring? (Enter a value from 1-5): ")
    impact = input(f"What is the potential impact of a {risk}? (Enter a value from 1-5): ")
    mitigation = input(f"What are the recommended mitigation strategies for a {risk}? ")

    # Append the risk data to the DataFrame
    risk_data = risk_data.append({'Risk': risk, 'Likelihood': likelihood, 'Impact': impact, 'Mitigation': mitigation}, ignore_index=True)

    # Create a pie chart of the likelihood and save it to a file
    fig, ax = plt.subplots()
    ax.pie([1, int(likelihood)], labels=['Low', 'High'], colors=['lightgray', 'dodgerblue'], autopct='%1.1f%%', startangle=90, counterclock=False)
    ax.axis('equal')
    ax.set_title(f"Likelihood of {risk}")
    plt.savefig(f"{risk}_likelihood.png")
    plt.clf()

    # Create a pie chart of the impact and save it to a file
    fig, ax = plt.subplots()
    ax.pie([1, int(impact)], labels=['Low', 'High'], colors=['lightgray', 'red'], autopct='%1.1f%%', startangle=90, counterclock=False)
    ax.axis('equal')
    ax.set_title(f"Impact of {risk}")
    plt.savefig(f"{risk}_impact.png")
    plt.clf()

# Create a new Excel workbook and write the risk data to a new sheet
wb = Workbook()
ws = wb.active
ws.title = "Cyber Security Risks"
for r in dataframe_to_rows(risk_data, index=False, header=True):
    ws.append(r)

# Save the workbook to a file
wb.save("PamperedPets_CyberSecurityRisks.xlsx")
# Create a bar chart of the likelihood and save it to a file
fig, ax = plt.subplots()
ax.bar(['Low', 'High'], [1, int(likelihood)], color=['lightgray', 'dodgerblue'])
ax.set_title(f"Likelihood of {risk}")
ax.set_ylabel('Frequency')
plt.savefig(f"{risk}_likelihood.png")
plt.clf()

# Create a bar chart of the impact and save it to a file
fig, ax = plt.subplots()
ax.bar(['Low', 'High'], [1, int(impact)], color=['lightgray', 'red'])
ax.set_title(f"Impact of {risk}")
ax.set_ylabel('Frequency')
plt.savefig(f"{risk}_impact.png")
plt.clf()