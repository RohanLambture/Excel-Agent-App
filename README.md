# Excel Data Analysis and Visualization App

## Overview

This Streamlit application allows users to upload an Excel file, query the data using natural language, and receive responses in the form of text or visual plots. The application leverages the LangChain Groq model for natural language processing and pandasai for smart dataframe operations.

## Features

- Upload and preview Excel files.
- Query the data using natural language.
- Generate visual plots based on the queries.
- Download the plot and data as an Excel file.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/RohanLambture/Excel-Agent-App.git
    cd Excel-Agent-App
    ```

2. Create a virtual environment and install the dependencies:

    ```bash
    python -m venv env
    source env/bin/activate  # On Windows, use `env\Scripts\activate`
    pip install -r requirements.txt
    ```

3. Set up the environment variables:

    Create a `.env` file in the project root directory and add your Groq API key:

    ```
    GROQ_API_KEY=your_groq_api_key
    ```

## Usage

1. Run the Streamlit app:

    ```bash
    streamlit run app.py
    ```

2. Open your web browser and go to `http://localhost:8501`.

3. Upload an Excel file, enter your query or plot prompt, and generate the response.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgements

This app uses the following libraries and services:

- [Streamlit](https://streamlit.io/)
- [LangChain Groq](https://groq.com/)
- [pandasai](https://github.com/pandasai/pandasai)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Repository

[Excel Data Analysis and Visualization App](https://github.com/RohanLambture/Excel-Agent-App)
