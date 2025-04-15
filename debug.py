import pdfplumber
import os

def convert_pdf_to_text(pdf_path, text_path=None):
    """
    Converts a PDF file to a text file.

    Args:
        pdf_path (str): The path to the PDF file.
        text_path (str, optional): The path to the output text file.
            If None, the text file will be saved in the same directory as the PDF
            with the same name but with a .txt extension. Defaults to None.

    Returns:
        str: The extracted text from the PDF, or None on error.
               Returns the path to the text file if text_path is provided.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() + "\n\n"  # Add page separator for readability

        if text_path:
            try:
                with open(text_path, 'w', encoding='utf-8') as f:
                    f.write(text)
                print(f"Text successfully extracted and saved to {text_path}")
                return text_path  # Return the path to the saved file
            except Exception as e:
                print(f"Error saving text to file: {e}")
                return None
        else:
            return text  # Return the extracted text
    except FileNotFoundError:
        print(f"Error: File not found at {pdf_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def main():
    """
    Main function to run the PDF to text converter.
    Prompts the user for the PDF file path and optional output path.
    """
    pdf_file = 'd2.pdf'  # Default PDF file name
    if not os.path.exists(pdf_file):
        print(f"Error: File not found at {pdf_file}")
        return

    save_option = input("Do you want to save the output to a text file? (y/n): ").lower()
    if save_option == 'y':
        text_file = input("Enter the path to save the text file (or press Enter to save in the same directory): ")
        if not text_file:
            # Generate a default text file name
            base_name = os.path.splitext(os.path.basename(pdf_file))[0]
            text_file = os.path.join(os.path.dirname(pdf_file), f"{base_name}.txt")
        # Ensure .txt extension is added
        if not text_file.lower().endswith(".txt"):
            text_file += ".txt"
        extracted_text_path = convert_pdf_to_text(pdf_file, text_file)
        if extracted_text_path:
            print(f"Text saved to: {extracted_text_path}")
        else:
            print("Failed to save the extracted text.")

    elif save_option == 'n':
        extracted_text = convert_pdf_to_text(pdf_file)
        if extracted_text:
            print("\nExtracted Text:\n")
            print(extracted_text)
        else:
            print("Failed to extract text from the PDF.")
    else:
        print("Invalid option. Please enter 'y' or 'n'.")

if __name__ == "__main__":
    main()

