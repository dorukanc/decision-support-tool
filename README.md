# Decision Support Tool built with Excel VBA

## Operation Research Term Project Decision Support Tool (DST)

**Overview**

Welcome to the **Operation Research Term Project Decision Support Tool (DST)** repository! This project showcases a decision support tool developed for operation research using **Excel VBA**. The DST is designed to facilitate queuing model calculations, ensuring accurate results while providing a user-friendly interface.

**Key Features**
- **User Interface:** Our DST utilizes **Excel VBA** to create a user-friendly interface for queuing model calculations.
- **Adaptability:** The tool adapts to different queuing models, allowing users to input parameters and receive calculated results.
- **Error Handling:** The DST incorporates error checks to alert users if they have not entered valid numbers or specific model selections.
- **Submission and Cleaning:** Two buttons are provided - one for submission and another for clearing the text box, enhancing user experience.
- **Alerts:** The tool provides alerts for scenarios such as selecting a queuing model before submission and entering positive values for Lambda.

**Getting Started**
1. Open the Excel sheet.
2. Launch the DS tool using the provided shortcut.
3. Enter the required parameters for the selected queuing model.
4. Click the "**Submit**" button to calculate results.
5. Use the "**Clear**" button to reset the text box.

**Error Handling**
- If no queuing model is selected before submission, the tool will display an alert: "**Please select a queuing model.**"
- If non-positive values are entered for Lambda, the tool will prompt: "**Please enter positive values for Lambda, Mu.**"

**Queuing Model Outputs**
- **M/M/1:GD/infinity/infinity:** Output details for the first queuing model.
- **M/M/1:(GD/N/∞):** Output details for the second queuing model.
- **M/M/c:(GD/∞/∞):** Expected answers for the third queuing model.
- **(M/M/c):(GD/N/∞):** Our answers for the (M/M/c):(GD/N/∞) queuing model.
- **(M/G/1):(GD/∞/∞):** Expected results for (M/G/1):(GD/∞/∞) queuing model.

**Conclusion**

Our project has successfully calculated results for all 7 different queuing models. The built-in error checks ensure that users cannot input non-valid values, enhancing the reliability of the tool.

**Developed by** Dorukan Çatak
