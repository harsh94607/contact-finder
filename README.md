# Apollo Contact Finder (web)

Find the right people at each company with the Apollo API, tick the ones you want,
reveal their emails and download everything as Excel.  *Developed by Harsh Goyal.*

## How it works
1. Type your Apollo API key in the sidebar (kept only in your browser session, never saved).
2. Upload an Excel file with two columns: `Company Name` and `Domain`.
3. Pick a company: the people Apollo finds are listed (free). Tick the ones you want,
   use **Load more** or **Search roles** for others.
4. **Reveal selected & save** uses Apollo credits (it asks first).
5. Download the Excel from the sidebar. Results live only in your browser session, so download before closing the tab.
   To continue later, upload that file again as "Previous output file".

## Deploy (free) on Streamlit Community Cloud
1. Create a new **private** GitHub repository and upload these files:
   `streamlit_app.py`, `requirements.txt`, `README.md`, `.gitignore`, and the folder `.streamlit/` (with `config.toml`).
   **Never upload your API key or any Excel file with contacts.**
2. Go to https://share.streamlit.io, sign in with GitHub, click **Create app**.
3. Choose the repository, branch `main`, main file `streamlit_app.py`, then **Deploy**.
4. Restrict who can open it: in the app, **Share** -> choose *Only specific people can view this app* and add your manager's email.

## Optional settings (App -> Settings -> Secrets)
```toml
APP_PASSWORD = "choose-a-password"     # asks everyone for this password first
APOLLO_API_KEY = "your-apollo-key"     # then nobody has to type a key (anyone who can open the app can spend your credits!)
```
Use `APOLLO_API_KEY` only together with `APP_PASSWORD` or a private app.

## Run locally (optional)
```
pip install -r requirements.txt
streamlit run streamlit_app.py
```
