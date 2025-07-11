# Project Delivery Dashboard

This Streamlit app allows project managers to:
- Add new projects
- Submit weekly updates
- View project reports with PDF export

## Run Locally

1. Install dependencies:
```
pip install -r requirements.txt
```

2. Create a `.streamlit/secrets.toml` file with your Azure SQL credentials.

3. Run the app:
```
streamlit run app.py
```

## Deploy to Streamlit Cloud

1. Push this repo to GitHub.
2. Go to https://share.streamlit.io and deploy the app using the repo.
