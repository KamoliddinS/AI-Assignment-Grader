
# AI Assignment Grader

A tool to automatically grade student assignments using OpenAI and send feedback in the form of PDFs via email.

## Features:

- Automatically grade student answers from Excel sheets.
- Generates a PDF report containing questions, student answers, scores, and feedback.
- Sends the generated PDF report via Gmail.

## Prerequisites:

Ensure you have Python 3.x installed. This script has been tested with Python 3.10, but it should work with other 3.x versions.

## Installation:

1. Clone this repository:

```bash
git clone AI-Assignment-Grader
cd AI-Assignment-Grader
```

2. Install the required packages:

```bash
pip install -r requirements.txt
```

3. Set up your environment variables in the `.env` file. An example is provided in the `.env.example` file. Copy it and fill in your actual data:

```bash
cp .env.example .env
nano .env
```

## Usage:

```bash
python assignment_grader.py path_to_excel_file.xlsx
```

Replace `path_to_excel_file.xlsx` with the path to your Excel file containing student answers.

## .env File Configuration:

Ensure you have the following keys set in your `.env` file:

```
OPEN_AI_API_KEY=YOUR_OPENAI_API_KEY
OPEN_AI_BASE_URL=YOUR_OPENAI_BASE_URL
OPEN_AI_API_TYPE=YOUR_API_TYPE
OPEN_AI_API_VERSION=YOUR_API_VERSION
OPEN_AI_DEPLOYMENT_ID=YOUR_DEPLOYMENT_ID
CREDENTIALS_JSON_PATH=YOUR_CREDENTIALS_JSON_PATH
```

## Contributing:

Contributions are welcome! Please read the contributing guidelines to get started.

## License:

[License name, e.g., MIT, GPL, etc.]

