# Configuration example

Do not store real credentials in GitHub.

## Google Apps Script

Use Script Properties:

- `GROQ_API_KEY` = your Groq API key

## Application configuration

Update only the non-secret settings in `Code.gs`:

- Crossref endpoint
- Groq model
- verification thresholds
- batch size / cache settings

Never commit the actual API key.
