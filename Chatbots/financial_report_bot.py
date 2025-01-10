# ==============================================================================
# Part 1 - Install necessary packages
# ==============================================================================
import os # For interacting with environment variables
from azure.storage.blob import BlobServiceClient #For interacting with Azure Blob Storage
import fitz #PyMuPDF, used to extract text from PDFs
from openai import AzureOpenAI # For interacting with Azure OpenAI GPT model's
import tkinter as tk # For creating a graphical user interface (GUI)
from tkinter import ttk, scrolledtext # Additional widgets for the GUI

# ==============================================================================
# Part 1 - Azure OpenAI Client Setup
# ==============================================================================
# Define the Azure OpenAI API key and endpoint, using environment variables to keep them secure
os.environ["AZURE_OPENAI_API_KEY"] = "add API key here"
os.environ["AZURE_OPENAI_ENDPOINT"] = "add endpoint here"

# Initialize the Azure OpenAI client to interact with the GPT-4 models
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"), # Fetches API key from the environment variable
    api_version="2023-05-15", # Version of  the API to use
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT") # Fetches the endpoint URL from the environment
)

# ==============================================================================
# Part 2 - Azure Blob Storage Setup
# ==============================================================================
# Define the connection string and container name to access Azure Blob Storage
connection_string = "DefaultEndpointsProtocol=https;AccountName=ban443;AccountKey=18YhXwIr5qy+DLuvNtvFAAS/lxRFhRx/z6yhKcr+n/E94ZunXBBLuo3XE7cVp/s0jC2usHKjFbDQ+AStWu6lNQ==;EndpointSuffix=core.windows.net"
# Specify the containter name
container_name ="chatbot-data"


# Create a BlobServiceClient object using the connection string
blob_service_client = BlobServiceClient.from_connection_string(connection_string)

# Create a Container Client to interact with the specified containter
container_client = blob_service_client.get_container_client(container_name)

# Print available reports (blobs) stored in the Azure container
print("Welcome to the this chatbot!")
print("The reports you can ask questions about are the following: ")
blobs=container_client.list_blobs()
for blob in blobs:
    print(" - ", blob.name) # Display each blob name(file name) that is available in the container

# ==============================================================================
# Part 3 - Downloading and Extracting Text from PDFs
# ==============================================================================
# Function to download a PDF file from Azure Blob Storage and extract its text content
def download_pdf(blob_name):
    """Downloads PDF file from Azure Blob Storage and extracts text"""
    blob_client = container_client.get_blob_client(blob_name) # Get the BlobClient for the file
    pdf_content=blob_client.download_blob().readall() # Download the blob's content

    #Save the PDF content to a local file
    pdf_path = f'downloaded_{blob_name}'
    with open(pdf_path, 'wb') as pdf_file:
        pdf_file.write(pdf_content)

    # Extract and return the text from the PDF file
    return extract_text_from_pdf(pdf_path)

# Function to extract text from a locally saved PDF file using PyMuPDF
def extract_text_from_pdf(pdf_path):
    """Extracts text from PDF file"""
    text = " "
    # Open the PDF file using PyMuPDF and extract text from each page
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text() # Append text from each page to the text variable
    return text # Return the complete extracted text

# ==============================================================================
# Part 4 - Function to Ask GPT-4 a Question
# ==============================================================================
# Function to ask GPT-4 a question based on the extracted PDF text and the user's question
def ask_gpt4(question, context):
    """Generates a response using GPT-4 based on a provided context and question."""
    # Create a prompt combining the PDF content and the user's question
    prompt = (
        f"The following is a summary of a financial year-end report:\n\n"
        f"{context}\n\n"
        f"Based on the above information, please answer the following question:\n"
        f"{question}"
    )

    try:
        # Send the prompt to Azure OpenAI GPT-4 and get a response
        response = client.chat.completions.create(
            model="GPT4o-API",  # GPT-4 deployment
            messages=[
                {"role": "system", "content": "You are a helpful assistant analyzing financial reports."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1000, # Limit the response length
            temperature=0.3 # Control randomness in the response
        )

        # Extract and return the response content
        content = response.choices[0].message.content
        return content # Return the resposne text

    except Exception as e:
        return f"Error with OpenAI API: {str(e)}" # Return the error message in case of failure

# ==============================================================================
# Part 5 - Main Chatbot GUI Class
# ==============================================================================
class ChatbotGUI:
    def __init__(self, master):
        """Initialize the GUI components and attach them to the master window"""
        self.master = master
        master.title("Financial Report Chatbot")  # Set the title of the window
        master.geometry("800x600")  # Set the window size

        # Allow the window to be resizable
        master.resizable(True, True)

        # Initialize instance attributes to store user input, blobs, and extracted text
        self.company_year_entry = None  # Input for the company names and years
        self.report_names = []  # List of available report names (from blobs)
        self.report_texts = {}  # Dictionary to store extracted text from the reports
        self.current_blobs = []  # Store the names of the reports being analyzed

        self.create_widgets()  # Create the GUI widgets
        self.load_blob_names()  # Load available report names from Blob Storage

    def create_widgets(self):
        """Create input fields and buttons for the chatbot interface"""
        # Input field for entering multiple company names and years
        ttk.Label(self.master, text="Enter at least one company name and year (e.g., Equinor 2020, Equinor 2021):").pack(pady=5)
        self.company_year_entry = ttk.Entry(self.master, width=70)
        self.company_year_entry.pack(pady=5)

        # Input field for entering the user's question
        ttk.Label(self.master, text="Enter your question:").pack(pady=5)
        self.question_entry = ttk.Entry(self.master, width=70)
        self.question_entry.pack(pady=5)

        # Create and display a button to submit the question
        ttk.Button(self.master, text="Ask", command=self.ask_question).pack(pady=10)

        # Scrollable text box for displaying the conversation
        self.chat_display = scrolledtext.ScrolledText(self.master, wrap=tk.WORD, width=80, height=20)
        self.chat_display.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

    def load_blob_names(self):
        """Load available reports from the blob container"""
        available_blobs = container_client.list_blobs()
        self.report_names = [blob.name for blob in available_blobs]

    def ask_question(self):
        """Handle the user's question and fetch the appropriate reports"""
        # Split the input by commas to handle multiple company-year pairs
        company_years = [entry.strip().lower().replace(" ", "-") + ".pdf" for entry in self.company_year_entry.get().split(',') if entry]
        question = self.question_entry.get()

        if not company_years or not question:
            self.chat_display.insert(tk.END, "Please enter at least one company name, year, and a question.\n\n")
            return

        matching_blobs = []
        for company_year in company_years:
            matching_blob = next((report_name for report_name in self.report_names if company_year in report_name.lower()), None)
            if matching_blob:
                matching_blobs.append(matching_blob)

        if not matching_blobs:
            self.chat_display.insert(tk.END, "No matching reports found.\n\n")
            return

        try:
            # Download and extract text from each matching report
            for blob in matching_blobs:
                if blob not in self.current_blobs:
                    self.report_texts[blob] = download_pdf(blob)
                    self.current_blobs.append(blob)
                    self.chat_display.insert(tk.END, f"Extracted text from {blob}...\n")

            # Analyze or compare multiple reports using GPT-4
            answer = self.analyze_reports(question, self.report_texts)
            self.chat_display.insert(tk.END, f"Analyzing reports...\n")
            self.chat_display.insert(tk.END, f"You: {question}\n")
            self.chat_display.insert(tk.END, f"Bot: {answer}\n\n")
        except Exception as e:
            self.chat_display.insert(tk.END, f"Error: {str(e)}\n\n")

        self.question_entry.delete(0, tk.END)
        self.chat_display.see(tk.END)

    @staticmethod
    def analyze_reports(question, report_texts):
        """Analyze multiple reports based on the provided question"""
        prompt = "The following are summaries of annual reports:\n\n"
        for i, (blob_name, text) in enumerate(report_texts.items(), 1):
            prompt += f"Report {i} ({blob_name}):\n{text}\n\n"

        prompt += f"Based on the above reports, please answer the following question:\n{question}"

        try:
            response = client.chat.completions.create(
                model="GPT4o-API",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant analyzing financial reports."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000,
                temperature=0.3
            )

            content = response.choices[0].message.content
            return content

        except Exception as e:
            return f"Error with OpenAI API: {str(e)}"

# ==============================================================================
# Part 6 - Main Function to Run the GUI
# ==============================================================================
def main():
    root = tk.Tk()  # Create the main window
    ChatbotGUI(root)  # Create an instance of ChatbotGUI, attaching it to the main window
    root.mainloop()  # Start the Tkinter event loop to display the window

# Run the main function when the script is executed
if __name__ == '__main__':
    main()
