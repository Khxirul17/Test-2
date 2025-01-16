import pandas as pd
import openai

# Step 1: Define the ExcelProcessor class
class ExcelProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None

    def read_excel(self):
        """Reads the Excel file and stores the data."""
        self.data = pd.read_excel(self.file_path)
        return self.data

    def process_data(self):
        """Processes the data to calculate average salary and department distribution."""
        if self.data is None:
            raise ValueError("No data found. Please read the Excel file first.")
        
        return {
            "average_salary": self.data["Salary"].mean(),
            "department_distribution": self.data["Department"].value_counts().to_dict()
        }

# Step 2: Define the GPT summarization function
def summarize_with_gpt(summary_data):
    """Generates a summary using OpenAI's GPT API."""
    prompt = (
        f"Summarize the following employee data:\n"
        f"Average Salary: {summary_data['average_salary']}\n"
        f"Department Distribution: {summary_data['department_distribution']}\n"
    )
    # Call GPT API
    response = openai.Completion.create(
        engine="text-davinci-003",  # Replace with your chosen GPT model
        prompt=prompt,
        max_tokens=100
    )
    return response.choices[0].text.strip()

# Main code
if __name__ == "__main__":
    # Initialize the processor with the provided Excel file
    file_path = "/path/to/your/Downloads/TableT2Q1.xlsx"  # Replace with your actual file path
    processor = ExcelProcessor(file_path)
    
    # Read and process the data
    data = processor.read_excel()
    summary_data = processor.process_data()

    # Summarize with GPT
    gpt_summary = summarize_with_gpt(summary_data)

    # Output the results
    print("Processed Data Summary:")
    print(summary_data)
    print("\nGPT-Generated Summary:")
    print(gpt_summary)
