import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlsxwriter
from datetime import datetime, timedelta

def get_user_input():
    """Prompt the user for input data."""
    print("Please provide the following data:")

    af_type = int(input("AF Type (1: Paroxysmal AF, 2: Persistent AF, 3: Long-Standing Persistent AF): "))
    symptom_severity = int(input("Symptom Severity (1: Asymptomatic/Mild, 2: Moderate, 3: Severe): "))
    response_to_therapy = int(input("Response to Pharmacological Therapy (1: No Issues, 2: ≤1 Drug, 3: ≥1 Drug): "))
    heart_failure = int(input("Presence of Heart Failure (1: No HF, 2: HFpEF, 3: HFrEF): "))
    patient_preference = int(input("Patient Preference (1: Neutral/Indifferent, 2: Strongly Prefers Ablation): "))
    risk_factors = int(input("Risk Factors and Co-Morbidities (1: No Significant, 2: Significant): "))
    thyroid_function = int(input("Thyroid Function Tests (0: Normal, 1: Dysfunction): "))
    renal_function = int(input("Renal Function Tests (0: Normal, 1: Impairment): "))
    electrolytes = int(input("Electrolytes (0: Normal, 1: Imbalance): "))
    inflammatory_markers = int(input("Inflammatory Markers (0: Normal, 1: Elevated CRP): "))
    bnp_ntprobn = int(input("BNP or NT-proBNP (0: Normal, 1: Elevated): "))
    coagulation_profile = int(input("Coagulation Profile (0: Normal, 1: Abnormal INR): "))
    cbc = int(input("Complete Blood Count (0: Normal, 1: Anemia/Thrombocytopenia): "))
    liver_function = int(input("Liver Function Tests (0: Normal, 1: Abnormal): "))
    age = int(input("Age (1: <60 Years, 2: 60-75 Years, 3: >75 Years): "))
    sex = int(input("Sex (1: Male, 2: Female): "))
    anticoagulant_use = int(input("Anticoagulant Use (0: Not on Anticoagulants, 1: On Anticoagulants): "))

    return (
        af_type, symptom_severity, response_to_therapy, heart_failure,
        patient_preference, risk_factors, thyroid_function, renal_function,
        electrolytes, inflammatory_markers, bnp_ntprobn, coagulation_profile,
        cbc, liver_function, age, sex, anticoagulant_use
    )

def calculate_af_score(
        af_type, symptom_severity, response_to_therapy, heart_failure,
        patient_preference, risk_factors, thyroid_function, renal_function,
        electrolytes, inflammatory_markers, bnp_ntprobn, coagulation_profile,
        cbc, liver_function, age, sex, anticoagulant_use):
    """Calculate the total score based on the provided criteria."""
    score = 0
    score += af_type
    score += symptom_severity
    score += response_to_therapy
    score += heart_failure
    score += patient_preference
    score += risk_factors
    score += thyroid_function
    score += renal_function
    score += electrolytes
    score += inflammatory_markers
    score += bnp_ntprobn
    score += coagulation_profile
    score += cbc
    score += liver_function
    score += age
    score += sex
    score += anticoagulant_use
    return score

def assess_ablation(score):
    """Assess whether ablation is recommended based on the total score."""
    if score < 13:
        return "Consider alternative treatments or further evaluation. Ablation may not be indicated at this time."
    elif 13 <= score <= 22:
        return "Ablation may be considered based on individual patient factors and preferences."
    else:
        return "Strong candidate for ablation. Proceed with evaluation for potential ablation therapy."

def predict_ablation_timing(score):
    """Predict how soon the patient needs to receive ablation based on the score."""
    if score < 13:
        return datetime.now() + timedelta(weeks=12)  # 12 weeks from now
    elif 13 <= score <= 22:
        return datetime.now() + timedelta(weeks=6)  # 6 weeks from now
    else:
        return datetime.now() + timedelta(weeks=2)  # 2 weeks from now

def create_number_line_chart(total_score):
    """Create a horizontal number line chart using Matplotlib."""
    fig, ax = plt.subplots(figsize=(10, 2))

    # Define the ranges and colors for the zones
    zones = [
        (0, 13, 'green', 'No Ablation'),
        (13, 23, 'yellow', 'Consider Ablation'),
        (23, 30, 'red', 'Needs Ablation')
    ]

    # Add colored zones to the chart
    for start, end, color, label in zones:
        ax.axhline(0, xmin=start/30, xmax=end/30, color=color, linewidth=8, label=label)

    # Plot the patient's score
    ax.plot(total_score, 0, 'bo', markersize=12)
    ax.text(total_score, 0.1, f'Patient Score: {total_score}', horizontalalignment='center', color='blue')

    # Customize the chart
    ax.set_xlim(0, 30)
    ax.set_yticks([])
    ax.set_xlabel('Score')
    ax.legend(loc='upper left')
    ax.set_title('Ablation Assessment Number Line')

    # Save the chart as an image file
    plt.savefig('number_line_chart.png')
    plt.close()

def create_excel_with_chart(total_score, recommendation, prediction_date):
    """Create an Excel file with the number line chart and other details."""
    # Create the number line chart
    create_number_line_chart(total_score)

    # Create a new Excel workbook and add a worksheet
    workbook = xlsxwriter.Workbook('Ablation_Assessment.xlsx')
    worksheet = workbook.add_worksheet('Ablation Assessment')

    # Write data to the Excel sheet
    worksheet.write_row('A1', ['Score', 'Label'])
    worksheet.write_row(1, 0, [0, 'No Ablation'])
    worksheet.write_row(2, 0, [13, 'Consider Ablation'])
    worksheet.write_row(3, 0, [23, 'Most Likely Needs Ablation'])
    worksheet.write_row(4, 0, [total_score, 'Patient Score'])

    # Add recommendation and prediction date to the Excel sheet
    recommendation_row = 6
    worksheet.write(recommendation_row, 0, "Recommendation:")
    worksheet.write(recommendation_row, 1, recommendation)
    worksheet.write(recommendation_row + 1, 0, "Prediction Date for Ablation:")
    worksheet.write(recommendation_row + 1, 1, prediction_date.strftime('%Y-%m-%d'))

    # Insert the chart image into the worksheet
    worksheet.insert_image('E2', 'number_line_chart.png')

    # Close the workbook
    workbook.close()

def main():
    # Get user input
    af_type, symptom_severity, response_to_therapy, heart_failure, \
    patient_preference, risk_factors, thyroid_function, renal_function, \
    electrolytes, inflammatory_markers, bnp_ntprobn, coagulation_profile, \
    cbc, liver_function, age, sex, anticoagulant_use = get_user_input()

    # Calculate total score
    total_score = calculate_af_score(
        af_type, symptom_severity, response_to_therapy, heart_failure,
        patient_preference, risk_factors, thyroid_function, renal_function,
        electrolytes, inflammatory_markers, bnp_ntprobn, coagulation_profile,
        cbc, liver_function, age, sex, anticoagulant_use
    )

    # Assess ablation recommendation
    recommendation = assess_ablation(total_score)

    # Predict ablation timing
    prediction_date = predict_ablation_timing(total_score)

    # Create Excel file with chart
    create_excel_with_chart(total_score, recommendation, prediction_date)

if __name__ == "__main__":
    main()
