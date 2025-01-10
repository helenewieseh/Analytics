# ==============================================================================
# Part 1 - Install packages
# ==============================================================================
import os # For managing environment variables
import requests # For making HTTP requests to external APIs (brreg)
from openai import AzureOpenAI # for interacting with Azure OpenAI GPT models
import csv # For writing CSV files
import tkinter as tk # For creating graphical user interface (GUI)
from tkinter import scrolledtext # For creating a scrollable text widget in the GUI

# ==============================================================================
# Part 2 - Azure OpenAI Client Setup
# ==============================================================================
# Set up environment variables for Azure OpenAI API key and endpoint
os.environ["AZURE_OPENAI_API_KEY"] = "add API key here"
os.environ["AZURE_OPENAI_ENDPOINT"] = "add endpoint here"

# Initialize the Azure OpenAI client with API key and endpoint
client = AzureOpenAI(
    api_key=os.getenv("AZURE_OPENAI_API_KEY"), # Retrieve API key from environment
    api_version="2023-05-15", # Specify the API version for compatibility
    azure_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT") # Set the API endpoint
)

# ==============================================================================
# Part 3 - BrregAPI Class for Brønnøysund Register Centre API
# ==============================================================================
class BrregAPI:
    BASE_URL = "https://data.brreg.no" # Base URL for Brreg API

    def __init__(self):
        # Initialize the necessary headers for API requests
        self.headers = {
            "Accept": "application/vnd.brreg.enhetsregisteret.enhet.v2+json"
        }

    # Method to fetch general services from the root API
    def fetch_services(self):
        """GET /api (Root) - Lists links to other services"""
        url = f"{self.BASE_URL}/enhetsregisteret/api"
        return self._get(url)

    # Method to search for entities by name
    def search_entities(self, name):
        """GET /api/enheter - Search for entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/enheter"
        params = {"navn": name}
        return self._get(url, params)

    # Method to fetch details for specific entities using organization number
    def fetch_entity(self, orgnr):
        """GET /api/enheter/{orgnr} - Fetch a specific entity"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/enheter/{orgnr}"
        return self._get(url)

    # Method to fetch roles for a specific entity based on its organization number
    def fetch_roles_for_entity(self, orgnr):
        """GET /api/enheter/{orgnr}/roller - Fetch all roles for a specific entity"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/enheter/{orgnr}/roller"
        return self._get(url)

    # method to fetch roles for a specific entity that requires personal ID
    def fetch_roles_with_personal_id(self, orgnr):
        """GET /autorisert-api/enheter/{orgnr}/roller - Fetch roles for a specific entity with personal ID"""
        url = f"{self.BASE_URL}/enhetsregisteret/autorisert-api/enheter/{orgnr}/roller"
        return self._get(url)

    # Method to download data of entities in different formats (e.g., JSON, CSV, XLSX)
    def download_entities(self, file_format="json"):
        """GET /api/enheter/lastned - Download entities in various formats (JSON, CSV, XLSX)"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/enheter/lastned"
        if file_format == "csv":
            url += "/csv" # If CSV format is required, append/CSV to the URL
        elif file_format == "xlsx":
            url += "/regneark" # If XLSX (spreadsheet) format is requested, append / regneark
        return self._get(url, stream=True) # Stream the response for large files

    # Method to search for sub-entities using Brreg API
    def search_sub_entities(self, name):
        """GET /api/underenheter - Search for sub-entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/underenheter"
        params = {"navn": name}
        return self._get(url, params)

    # Method to fetch details of a specific sub-entity using its organization number
    def fetch_sub_entity(self, orgnr):
        """GET /api/underenheter/{orgnr} - Fetch a specific sub-entity"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/underenheter/{orgnr}"
        return self._get(url)

    # Method to download data of sub-entities in various formats
    def download_sub_entities(self, file_format="json"):
        """GET /api/underenheter/lastned - Download sub-entities in various formats (JSON, CSV, XLSX)"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/underenheter/lastned"
        if file_format == "csv":
            url += "/csv"
        elif file_format == "xlsx":
            url += "/regneark"
        return self._get(url, stream=True)

    # Method to fetch updates for entities
    def fetch_entity_updates(self):
        """GET /api/oppdateringer/enheter - Fetch updated entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/oppdateringer/enheter"
        return self._get(url)

    # Method to fetch updates for sub-entities
    def fetch_sub_entity_updates(self):
        """GET /api/oppdateringer/underenheter - Fetch updated sub-entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/oppdateringer/underenheter"
        return self._get(url)

    # Method to fetch updates of roles in entities
    def fetch_role_updates(self):
        """GET /api/oppdateringer/roller - Fetch role updates"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/oppdateringer/roller"
        return self._get(url)

    # Method to fetch all organization forms available in Brreg
    def fetch_org_forms(self):
        """GET /api/organisasjonsformer - Fetch all forms of organisation"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/organisasjonsformer"
        return self._get(url)

    # Method to fetch organization forms for entities
    def fetch_org_form_for_entities(self):
        """GET /api/organisasjonsformer/enheter - Fetch forms of organisation for entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/organisasjonsformer/enheter"
        return self._get(url)

    # Method to fetch organization forms for sub-entities
    def fetch_org_form_for_sub_entities(self):
        """GET /api/organisasjonsformer/underenheter - Fetch forms of organisation for sub-entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/organisasjonsformer/underenheter"
        return self._get(url)

    # Method to fetch a specific organization form based on its code
    def fetch_org_form(self, org_code):
        """GET /api/organisasjonsformer/{orgkode} - Fetch a given form of organisation"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/organisasjonsformer/{org_code}"
        return self._get(url)

    # Method to download the total inventory of roles for all entities
    def download_total_roles(self):
        """GET /api/roller/totalbestand - Download total inventory of roles for all entities"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/roller/totalbestand"
        return self._get(url)

    # Method to fetch the different types of roles in the brreg API
    def fetch_role_types(self):
        """GET /api/roller/rolletyper - Fetch role types"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/roller/rolletyper"
        return self._get(url)

    # Method to fetch grou types of roles
    def fetch_role_group_types(self):
        """GET /api/roller/rollegruppetyper - Fetch role group types"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/roller/rollegruppetyper"
        return self._get(url)

    # Method to fetch representatives related to roles
    def fetch_representatives(self):
        """GET /api/roller/representanter - Fetch representatives"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/roller/representanter"
        return self._get(url)

    # Method to fetch cadastral unit information
    def fetch_cadastral_unit(self, unit_id):
        """GET /api/matrikkelenhet - Fetch a given cadastral unit"""
        url = f"{self.BASE_URL}/enhetsregisteret/api/matrikkelenhet"
        params = {"id": unit_id}
        return self._get(url, params)

    # Method to download CSV for all political parties registered
    def download_political_parties_csv(self):
        """GET /partiregisteret/api/lastned/csv - Download total inventory of political parties in CSV format"""
        url = f"{self.BASE_URL}/partiregisteret/api/lastned/csv"
        return self._get(url, stream=True)

    # Method to download CSV of all non-profit organizations registered
    def download_non_profit_orgs_csv(self):
        """GET /frivillighetsregisteret/api/lastned/csv - Download total inventory of non-profit organisations in CSV format"""
        url = f"{self.BASE_URL}/frivillighetsregisteret/api/lastned/csv"
        return self._get(url, stream=True)

    # Helper method for making GET requests and handling responses
    def _get(self, url, params=None, stream=False):
        """Helper method for making GET requests"""
        response = requests.get(url, headers=self.headers, params=params, stream=stream)
        if response.status_code == 200:
            # Return JSON resposne for non-streaming content or content for streamed responses
            return response.json() if not stream else response.content
        else:
            # Print error message and return None in case of failure
            print(f"Error: {response.status_code}")
            return None

    @staticmethod
    def extract_data(json_data):
        if not json_data:
            return []

        entities = json_data['_embedded']['enheter']
        extracted_data = []
        for entity in entities:
            data = {
                'organisasjonsnummer': entity.get('organisasjonsnummer', ''),
                'navn': entity.get('navn', ''),
                'organisasjonsform': entity.get('organisasjonsform', {}).get('beskrivelse', ''),
                'registreringsdatoEnhetsregisteret': entity.get('registreringsdatoEnhetsregisteret', ''),
                'naeringskode1': entity.get('naeringskode1', {}).get('beskrivelse', ''),
                'forretningsadresse': ', '.join(entity.get('forretningsadresse', {}).get('adresse', [])),
                'postnummer': entity.get('forretningsadresse', {}).get('postnummer', ''),
                'poststed': entity.get('forretningsadresse', {}).get('poststed', ''),
                'kommune': entity.get('forretningsadresse', {}).get('kommune', ''),
                'kommunenummer': entity.get('forretningsadresse', {}).get('kommunenummer', ''),
                'epostadresse': entity.get('epostadresse', ''),
                'telefon': entity.get('telefon', ''),
                'hjemmeside': entity.get('hjemmeside', ''),
                'stiftelsesdato': entity.get('stiftelsesdato', ''),
                'sisteInnsendteAarsregnskap': entity.get('sisteInnsendteAarsregnskap', ''),
            }
            extracted_data.append(data)
        return extracted_data

    # Save data to CSV file
    def save_to_csv(self, data,
                    filename='/Users/helenewiese-hansen/Library/CloudStorage/OneDrive-NorwegianSchoolofEconomics/BAN443/pythonProject/.venv/bin/python /Users/helenewiese-hansen/Downloads/financial_report_bot.csv'):
        if not data:
            print("No data available to save.")
            return

        keys = data[0].keys()

        # Open the file with UTF-8 with BOM to ensure Excel handles encoding correctly
        with open(filename, 'w', newline='', encoding='utf-8-sig') as output_file:
            dict_writer = csv.DictWriter(output_file, fieldnames=keys, quoting=csv.QUOTE_MINIMAL)
            dict_writer.writeheader()
            dict_writer.writerows(data)

        print(f"Data saved to {filename}")


# ==============================================================================
# Part 4 - Function to ask Azure OpenAI a question based on brreg data
# ==============================================================================
def ask_azure_openai(question, brreg_data):
    """
    This function sends a user question along with brreg data as context to Azure OpenAI
    and returns a response.
    """
    # Prepare the prompt for OpenAI by combining Brreg data with the user question
    prompt = f"The following is data from the Brønnøysund Register Centre:\n\n{brreg_data}\n\nAnswer the following question: {question}"

    try:
        # Make a request using the AzureOpenAI client
        response = client.chat.completions.create(
            model="GPT4o-API",  # Deployment name
            messages=[
                {"role": "system", "content": "You are a helpful assistant analyzing financial reports."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1000, # Limit the responses to 1000 tokens
            temperature=0.3 # Adjust the creativity level of the response
        )

        # Extract and return the response content
        content = response.choices[0].message.content
        return content

    except Exception as e:
        # Handle errors during the API request
        return f"Error with OpenAI API: {str(e)}"

# ==============================================================================
# Part 5 - Function to handle the chat interaction
# ==============================================================================
def chat_interaction():
    """
    Handles the interaction between the user and the chatbot. Retrieves the user's question
    and calls the appropriate function to get data from Brreg and OpenAI.
    """
    # Retrieve the user's question from the input field
    question = question_entry.get()

    # Check if the user has entered a question
    if not question:
        chat_display.insert(tk.END, "Bot: Please enter a question.\n\n")
        return

    # Fetching data for the first time, if not already fetched
    if not hasattr(chat_interaction, "brreg_data"):
        entity_name = entity_entry.get()
        if not entity_name:
            chat_display.insert(tk.END, "Bot: Please enter an entity name.\n\n")
            return

        brreg_api = BrregAPI()
        json_data = brreg_api.search_entities(entity_name)

        if not json_data:
            chat_display.insert(tk.END, f"Bot: Could not retrieve data for {entity_name} from brreg.\n\n")
            return

        extracted_data = brreg_api.extract_data(json_data)
        chat_interaction.brreg_data = extracted_data

        # Save to CSV if the question involves "download CSV"
        if "download" in question.lower() and "csv" in question.lower():
            brreg_api.save_to_csv(extracted_data, filename=f"{entity_name}_entities_clean.csv")
            chat_display.insert(tk.END,
                                f"Bot: Data for {entity_name} has been downloaded in clean CSV format for Excel.\n\n")
            return

    # Send the user's question along with the Brreg data to Azure OpenAI for analysis
    answer = ask_azure_openai(question, chat_interaction.brreg_data)

    # Display the user's question and the chatbot's answer in the chat display
    chat_display.insert(tk.END, f"You: {question}\n")
    chat_display.insert(tk.END, f"Bot: {answer}\n\n")

    # Clear the question input field for the next question
    question_entry.delete(0, tk.END)

    # Scroll to the bottom of the chat display to show the latest question
    chat_display.see(tk.END)

# ==============================================================================
# Part 6 - GUI Setup using Tkinter
# ==============================================================================
# Create the main window for the chatbot
root = tk.Tk()
root.title("Brreg Chatbot") # Set the title of the window
root.geometry("600x400") # Set the dimensions of the window

# Create the labels and input fields for entity name and questions
entity_label = tk.Label(root, text="Entity Name:")
entity_label.pack() # Display the label on the window
entity_entry = tk.Entry(root, width=50) # Create an entry field for the entity name
entity_entry.pack() # Add the entry field to the window

question_label = tk.Label(root, text="Question:")
question_label.pack() # Display the label on the window
question_entry = tk.Entry(root, width=50) # Create an entry field for the question
question_entry.pack() # Add the entry field to the window

# Create the "Ask" button that will trigger the chat interaction
chat_button = tk.Button(root, text="Ask", command=chat_interaction)
chat_button.pack() # Display the button on the window

# Create a scrollable text box to display the conversation
chat_display = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=70, height=20)
chat_display.pack(padx=10, pady=10) # Add padding around the text box for a clean layout

# Start the Tkinter main even loop, which keeps the GUI running
root.mainloop()
