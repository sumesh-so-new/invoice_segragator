import streamlit as st
import os, json, sys, tempfile, shutil, re
from pathlib import Path
from collections import defaultdict
import pandas as pd

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Invoice & Credit Note Processor",
    page_icon="📄",
    layout="wide",
)

# ── Import main6 functions ────────────────────────────────────────────────────
sys.path.insert(0, os.path.dirname(__file__))
from main6 import (
    process_folder,
    is_credit_note,
    _detect_brand,
    _BRAND_KEYWORDS,
)

# ── Google OAuth libs ─────────────────────────────────────────────────────────
# pip install google-auth google-auth-oauthlib google-api-python-client
try:
    from google_auth_oauthlib.flow import Flow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2.credentials import Credentials
    import io
    GOOGLE_LIBS_OK = True
except ImportError:
    GOOGLE_LIBS_OK = False

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title   { font-size: 2rem; font-weight: 700; color: #1F4E79; }
    .sub-title    { font-size: 1rem; color: #555; margin-bottom: 1.5rem; }
    .stat-box     { background: #f0f4fb; border-radius: 10px; padding: 16px 20px;
                    border-left: 4px solid #1F4E79; }
    .stat-num     { font-size: 2rem; font-weight: 700; color: #1F4E79; }
    .stat-label   { font-size: 0.85rem; color: #555; }
    .brand-pill   { display: inline-block; background: #e8f0fe; color: #1a56db;
                    border-radius: 20px; padding: 3px 12px; margin: 3px;
                    font-size: 0.82rem; font-weight: 600; }
    .success-box  { background: #ecfdf5; border: 1px solid #6ee7b7;
                    border-radius: 8px; padding: 14px 18px; color: #065f46; }
    .file-row     { background: #f8fafc; border-radius: 6px;
                    padding: 8px 14px; margin: 4px 0; font-size: 0.88rem; }
    .inv-tag      { color: #1d4ed8; font-weight: 600; }
    .cn-tag       { color: #7c3aed; font-weight: 600; }
    .user-badge   { background: #e8f5e9; border-radius: 20px; padding: 6px 14px;
                    font-size: 0.9rem; color: #2e7d32; font-weight: 500; }
    .google-btn   { background: #fff; border: 2px solid #dadce0; border-radius: 8px;
                    padding: 12px 28px; font-size: 1rem; font-weight: 500;
                    color: #3c4043; cursor: pointer; display: inline-flex;
                    align-items: center; gap: 10px; text-decoration: none; }
    .google-btn:hover { border-color: #1F4E79; box-shadow: 0 2px 8px rgba(0,0,0,0.12); }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# GOOGLE OAUTH HELPERS
# ══════════════════════════════════════════════════════════════════════════════

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/drive.readonly",
]


def _get_client_config():
    """
    Read OAuth credentials.
    Priority: environment variables → .streamlit/secrets.toml fallback.

    Set these in your environment (or .env loaded via python-dotenv):
        GOOGLE_CLIENT_ID     = "....apps.googleusercontent.com"
        GOOGLE_CLIENT_SECRET = "GOCSPX-..."
        NEXTAUTH_URL         = "http://localhost:8502"   # redirect URI
    """
    client_id     = os.environ.get("GOOGLE_CLIENT_ID", "")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET", "")
    redirect_uri  = os.environ.get("NEXTAUTH_URL", "")

    # Fallback to secrets.toml if env vars not set
    if not client_id or not client_secret:
        try:
            s = st.secrets.get("google_oauth", {})
            client_id     = client_id     or s.get("client_id", "")
            client_secret = client_secret or s.get("client_secret", "")
            redirect_uri  = redirect_uri  or s.get("redirect_uri", "http://localhost:8502")
        except Exception:
            redirect_uri  = redirect_uri  or "http://localhost:8502"

    return {
        "web": {
            "client_id":     client_id,
            "client_secret": client_secret,
            "redirect_uris": [redirect_uri],
            "auth_uri":      "https://accounts.google.com/o/oauth2/auth",
            "token_uri":     "https://oauth2.googleapis.com/token",
        }
    }


def _build_flow():
    cfg      = _get_client_config()
    redirect = cfg["web"]["redirect_uris"][0]
    flow     = Flow.from_client_config(cfg, scopes=SCOPES, redirect_uri=redirect)
    # Disable PKCE — prevents "Missing code verifier" error
    flow.oauth2session._client.code_challenge_method = None
    if hasattr(flow.oauth2session, "_code_challenge_method"):
        flow.oauth2session._code_challenge_method = None
    return flow


def _get_auth_url():
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"]  = "1"

    flow = _build_flow()
    auth_url, state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )

    # Strip code_challenge params if they snuck in
    from urllib.parse import urlparse, urlencode, parse_qs
    parsed = urlparse(auth_url)
    params = parse_qs(parsed.query, keep_blank_values=True)
    params.pop("code_challenge",        None)
    params.pop("code_challenge_method", None)
    clean_query = urlencode({k: v[0] for k, v in params.items()})
    auth_url = parsed._replace(query=clean_query).geturl()

    st.session_state["oauth_state"] = state
    return auth_url


def _exchange_code(code):
    """Exchange auth code — direct POST to token endpoint, no PKCE."""
    import requests as _requests
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"]  = "1"

    cfg           = _get_client_config()
    redirect_uri  = cfg["web"]["redirect_uris"][0]
    client_id     = cfg["web"]["client_id"]
    client_secret = cfg["web"]["client_secret"]
    token_uri     = cfg["web"]["token_uri"]

    resp = _requests.post(
        token_uri,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        data={
            "code":          code,
            "client_id":     client_id,
            "client_secret": client_secret,
            "redirect_uri":  redirect_uri,
            "grant_type":    "authorization_code",
        },
    )

    if not resp.ok:
        raise RuntimeError(f"Token exchange failed {resp.status_code}: {resp.text}")

    token_data = resp.json()
    if "error" in token_data:
        raise RuntimeError(f"Token error: {token_data}")

    st.session_state["google_creds"] = {
        "token":         token_data["access_token"],
        "refresh_token": token_data.get("refresh_token"),
        "token_uri":     token_uri,
        "client_id":     client_id,
        "client_secret": client_secret,
        "scopes":        SCOPES,
    }


def _get_creds():
    raw = st.session_state.get("google_creds")
    if not raw:
        return None
    return Credentials(
        token=raw["token"],
        refresh_token=raw["refresh_token"],
        token_uri=raw["token_uri"],
        client_id=raw["client_id"],
        client_secret=raw["client_secret"],
        scopes=raw["scopes"],
    )


def _get_user_info(creds):
    svc = build("oauth2", "v2", credentials=creds)
    return svc.userinfo().get().execute()


def _list_drive_folders(creds, parent_id="root"):
    svc = build("drive", "v3", credentials=creds)
    q   = (f"'{parent_id}' in parents "
           "and mimeType='application/vnd.google-apps.folder' "
           "and trashed=false")
    res = svc.files().list(q=q, fields="files(id, name)", orderBy="name").execute()
    return res.get("files", [])


def _list_pdfs_in_folder(creds, folder_id):
    svc = build("drive", "v3", credentials=creds)
    q   = (f"'{folder_id}' in parents "
           "and mimeType='application/pdf' "
           "and trashed=false")
    res = svc.files().list(
        q=q, fields="files(id, name, size)", orderBy="name"
    ).execute()
    return res.get("files", [])


def _download_pdfs_to_temp(creds, pdf_files):
    """Download all PDFs into a local temp directory. Returns the temp path."""
    svc     = build("drive", "v3", credentials=creds)
    tmp_dir = tempfile.mkdtemp(prefix="drive_pdfs_")
    prog    = st.progress(0)
    status  = st.empty()

    for i, f in enumerate(pdf_files, 1):
        status.markdown(f"⬇️ Downloading **{f['name']}** ({i}/{len(pdf_files)})…")
        request    = svc.files().get_media(fileId=f["id"])
        buf        = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        with open(os.path.join(tmp_dir, f["name"]), "wb") as out:
            out.write(buf.getvalue())
        prog.progress(int(i / len(pdf_files) * 100))

    status.empty()
    prog.empty()
    return tmp_dir


def _extract_folder_id_from_link(link: str):
    """
    Extract a Google Drive folder ID from various URL formats:
      https://drive.google.com/drive/folders/FOLDER_ID
      https://drive.google.com/drive/folders/FOLDER_ID?usp=sharing
      https://drive.google.com/open?id=FOLDER_ID
      https://drive.google.com/folderview?id=FOLDER_ID
    Returns folder_id string or None.
    """
    if not link:
        return None
    # /folders/FOLDER_ID
    m = re.search(r'/folders/([a-zA-Z0-9_-]+)', link)
    if m:
        return m.group(1)
    # ?id=FOLDER_ID or &id=FOLDER_ID
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', link)
    if m:
        return m.group(1)
    return None


# ══════════════════════════════════════════════════════════════════════════════
# SHARED PROCESSING + RESULTS
# ══════════════════════════════════════════════════════════════════════════════

def run_processing(folder_path):
    """Run process_folder and return (final_data, xlsx_file, invoices, credit_notes)."""
    st.subheader("⚙️ Processing PDFs…")
    log_area  = st.empty()
    progress  = st.progress(0)
    log_lines = []

    def log(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines), language=None)

    import builtins
    _orig = builtins.print
    def _st_print(*args, **kwargs):
        log(" ".join(str(a) for a in args))
        _orig(*args, **kwargs)
    builtins.print = _st_print

    try:
        progress.progress(10)
        invoices, credit_notes = process_folder(
            folder_path, os.path.join(folder_path, "output.json")
        )
        progress.progress(100)
    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)
        builtins.print = _orig
        return None
    finally:
        builtins.print = _orig

    final_file = os.path.join(folder_path, "final.json")
    xlsx_file  = os.path.join(folder_path, "final.xlsx")

    with open(final_file, encoding="utf-8") as f:
        final_data = json.load(f)

    return final_data, xlsx_file, invoices, credit_notes


def show_results(final_data, xlsx_file, invoices, credit_notes):
    matched   = sum(1 for r in final_data if r.get("Credit Note No") is not None)
    unmatched = len(final_data) - matched

    brand_counts = defaultdict(int)
    for row in final_data:
        brand_counts[_detect_brand(row.get("Narration", ""))] += 1

    st.divider()
    st.subheader("✅ Processing Complete")

    m1, m2, m3, m4 = st.columns(4)
    for col, num, label in [
        (m1, len(invoices),     "Invoice Items"),
        (m2, len(credit_notes), "Credit Note Items"),
        (m3, matched,           "Matched (ASN found)"),
        (m4, unmatched,         "Unmatched (null)"),
    ]:
        col.markdown(
            f"<div class='stat-box'><div class='stat-num'>{num}</div>"
            f"<div class='stat-label'>{label}</div></div>",
            unsafe_allow_html=True,
        )

    # Brand pills
    st.markdown("#### 🏷️ Brand Breakdown")
    st.markdown(
        "".join(
            f"<span class='brand-pill'>{b} ({c})</span>"
            for b, c in sorted(brand_counts.items())
        ),
        unsafe_allow_html=True,
    )

    # Download button
    st.markdown("#### 📊 Download Excel")
    if os.path.exists(xlsx_file):
        with open(xlsx_file, "rb") as xf:
            st.download_button(
                label="⬇️  Download final.xlsx  (Brand-wise Sheets)",
                data=xf.read(),
                file_name="final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        size_kb = os.path.getsize(xlsx_file) / 1024
        st.caption(f"📊 {len(brand_counts)} brand sheets · {size_kb:.1f} KB")

    # Preview tabs
    st.markdown("#### 👁️ Data Preview")
    df_all = pd.DataFrame(final_data)
    tabs   = st.tabs(["All Records"] + sorted(brand_counts.keys()))
    with tabs[0]:
        st.dataframe(df_all, use_container_width=True, height=400)
    for tab, brand in zip(tabs[1:], sorted(brand_counts.keys())):
        with tab:
            df_b = df_all[
                df_all["Narration"].apply(lambda n: _detect_brand(n) == brand)
            ].reset_index(drop=True)
            st.dataframe(df_b, use_container_width=True, height=400)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN UI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown('<div class="main-title">📄 Invoice & Credit Note Processor</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Process PDF invoices and credit notes, match by ASN, and export brand-wise Excel sheets.</div>', unsafe_allow_html=True)
st.divider()

# ── Handle OAuth callback (Google redirects back with ?code=...) ─────────────
if GOOGLE_LIBS_OK:
    params = st.query_params
    if "code" in params and "google_creds" not in st.session_state:
        with st.spinner("Completing Google sign-in…"):
            try:
                os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
                os.environ["OAUTHLIB_RELAX_TOKEN_SCOPE"]  = "1"
                _exchange_code(params["code"])
                st.query_params.clear()
                st.rerun()
            except Exception as e:
                st.error(f"OAuth error: {e}")
                st.error("Full error details:")
                import traceback
                st.code(traceback.format_exc())

# ── Source selector tabs ──────────────────────────────────────────────────────
tab_local, tab_drive = st.tabs(["📁  Local Folder", "☁️  Google Drive"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — LOCAL FOLDER
# ══════════════════════════════════════════════════════════════════════════════
with tab_local:
    st.markdown("#### Enter the path to your PDFs folder")

    # ── Folder picker via tkinter (runs on user's local machine) ─────────────
    def _pick_folder() -> str:
        """Open a native OS folder-picker dialog and return the chosen path."""
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            # Windows: force dialog to foreground
            root.wm_attributes("-topmost", True)
            root.after(100, root.focus_force)
            selected = filedialog.askdirectory(
                parent=root,
                title="Select folder containing PDFs"
            )
            root.destroy()
            return selected or ""
        except Exception as e:
            st.error(f"❌ Could not open folder picker: {e}")
            return ""

    # KEY FIX: pre-initialise the widget's session_state key so Browse can
    # write directly into it and st.rerun() will show the updated value.
    if "local_folder" not in st.session_state:
        st.session_state["local_folder"] = ""

    col_path, col_browse, col_btn = st.columns([3, 1, 1])

    with col_browse:
        browse_clicked = st.button("📂 Browse", use_container_width=True, key="browse_folder")

    with col_btn:
        run_local = st.button("▶ Process", type="primary",
                              use_container_width=True, key="run_local")

    # Handle Browse BEFORE rendering the text_input so the picked path
    # is already in session_state when the widget is drawn.
    if browse_clicked:
        picked = _pick_folder()
        if picked:
            # Write directly into the widget's own key → shows up immediately
            st.session_state["local_folder"] = picked
        else:
            st.warning("⚠️ No folder selected. You can also type the path manually.")

    with col_path:
        folder_path = st.text_input(
            "Folder path",
            placeholder=r"e.g. D:\Sumesh\invoice_data_extractor\PUMA\New folder",
            label_visibility="collapsed",
            key="local_folder",   # reads & writes st.session_state["local_folder"]
        )

    if folder_path and os.path.isdir(folder_path):
        pdf_files = sorted(
            [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        )
        if pdf_files:
            inv_prev = [f for f in pdf_files if not is_credit_note(f)]
            cn_prev  = [f for f in pdf_files if is_credit_note(f)]
            with st.expander(f"📂 Found {len(pdf_files)} PDFs — click to preview"):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**<span class='inv-tag'>📑 Invoices ({len(inv_prev)})</span>**",
                                unsafe_allow_html=True)
                    for f in inv_prev:
                        st.markdown(f"<div class='file-row'>📄 {f}</div>",
                                    unsafe_allow_html=True)
                with c2:
                    st.markdown(f"**<span class='cn-tag'>🔖 Credit Notes ({len(cn_prev)})</span>**",
                                unsafe_allow_html=True)
                    for f in cn_prev:
                        st.markdown(f"<div class='file-row'>📄 {f}</div>",
                                    unsafe_allow_html=True)
        else:
            st.warning("⚠️ No PDF files found in this folder.")
    elif folder_path:
        st.error("❌ Folder path does not exist.")

    if run_local:
        if not folder_path or not os.path.isdir(folder_path):
            st.error("Please enter a valid folder path.")
        else:
            result = run_processing(folder_path)
            if result:
                show_results(*result)
    elif not folder_path:
        st.info("👆 Enter a folder path above and click **▶ Process** to begin.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — GOOGLE DRIVE
# ══════════════════════════════════════════════════════════════════════════════
with tab_drive:

    if not GOOGLE_LIBS_OK:
        st.error("Google libraries not installed. Run:\n```\npip install google-auth google-auth-oauthlib google-api-python-client\n```")
        st.stop()

    # Check secrets are configured
    client_id     = os.environ.get("GOOGLE_CLIENT_ID", "")
    client_secret = os.environ.get("GOOGLE_CLIENT_SECRET", "")
    if not client_id or not client_secret:
        try:
            _go = st.secrets.get("google_oauth", {})
            client_id     = client_id     or _go.get("client_id", "")
            client_secret = client_secret or _go.get("client_secret", "")
        except Exception:
            pass
    secrets_ok = bool(client_id and client_secret)

    if not secrets_ok:
        st.warning("⚙️ **Google OAuth not configured.**")

        st.markdown("#### Option 1 — Environment variables (recommended)")
        st.code("""
# .env file
GOOGLE_CLIENT_ID     = "YOUR_CLIENT_ID.apps.googleusercontent.com"
GOOGLE_CLIENT_SECRET = "YOUR_CLIENT_SECRET"
NEXTAUTH_URL         = "http://localhost:8502"
        """, language="bash")

        st.markdown("#### Option 2 — `.streamlit/secrets.toml`")
        st.code("""
[google_oauth]
client_id     = "YOUR_CLIENT_ID.apps.googleusercontent.com"
client_secret = "YOUR_CLIENT_SECRET"
redirect_uri  = "http://localhost:8502"
        """, language="toml")

        with st.expander("📋 How to get Google OAuth credentials — step by step"):
            st.markdown("""
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. **Create a project** (or select existing one)
3. Enable **Google Drive API** → APIs & Services → Library → search "Drive"
4. Enable **Google People API** (for user info)
5. **APIs & Services → OAuth consent screen** → External → fill App name & email
6. **Credentials → Create OAuth 2.0 Client ID** → Web application
7. Under *Authorized redirect URIs* add: `http://localhost:8502`
8. Copy **Client ID** and **Client Secret** into `.env` or `secrets.toml`
9. Restart the Streamlit app
            """)

    else:
        # ── Not logged in ──────────────────────────────────────────────────────
        if "google_creds" not in st.session_state:
            st.markdown("### 🔐 Sign in with Google")
            st.markdown(
                "Connect your Google account to browse Drive folders "
                "and pick the folder containing your invoice PDFs."
            )
            auth_url = _get_auth_url()
            st.link_button(
                "🔑  Sign in with Google",
                auth_url,
                use_container_width=False,
                type="primary",
            )

        # ── Logged in ──────────────────────────────────────────────────────────
        else:
            creds = _get_creds()

            # Load user info once
            if "google_user" not in st.session_state:
                with st.spinner("Loading account info…"):
                    st.session_state["google_user"] = _get_user_info(creds)

            user = st.session_state["google_user"]

            # User badge + logout
            ucol, lcol = st.columns([6, 1])
            with ucol:
                st.markdown(
                    f"<span class='user-badge'>✅ &nbsp;"
                    f"{user.get('name', '')} &nbsp;·&nbsp; "
                    f"{user.get('email', '')}</span>",
                    unsafe_allow_html=True,
                )
            with lcol:
                if st.button("🚪 Logout", use_container_width=True):
                    for key in ["google_creds", "google_user", "drive_nav"]:
                        st.session_state.pop(key, None)
                    st.rerun()

            st.markdown("")
            st.markdown("### ☁️ Browse & Select a Folder")

            # ── Paste a Google Drive folder link ──────────────────────────────
            st.markdown("🔗 **Paste a Google Drive folder link**")
            link_col, btn_col = st.columns([5, 1])
            with link_col:
                folder_link = st.text_input(
                    "Google Drive folder URL",
                    placeholder="https://drive.google.com/drive/folders/1ABC...",
                    label_visibility="collapsed",
                    key="drive_folder_link",
                )
            with btn_col:
                go_link = st.button(
                    "📂 Go",
                    key="go_link",
                    use_container_width=True,
                    type="primary",
                )

            if go_link:
                folder_id = _extract_folder_id_from_link(folder_link)
                if folder_id:
                    try:
                        svc         = build("drive", "v3", credentials=creds)
                        meta        = svc.files().get(
                            fileId=folder_id, fields="id, name"
                        ).execute()
                        folder_name = meta.get("name", folder_id)
                        st.session_state["drive_nav"] = [
                            ("root", "My Drive"),
                            (folder_id, folder_name),
                        ]
                        st.success(f"✅ Navigated to **{folder_name}**")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Could not access folder: {e}")
                else:
                    st.warning(
                        "⚠️ Could not find a folder ID in that link. "
                        "Please paste a valid Google Drive folder URL."
                    )
            st.markdown("---")

            # Breadcrumb navigation stack: [(folder_id, folder_name), ...]
            if "drive_nav" not in st.session_state:
                st.session_state["drive_nav"] = [("root", "My Drive")]

            nav = st.session_state["drive_nav"]

            # ── Breadcrumb bar ─────────────────────────────────────────────────
            crumb_cols = st.columns(len(nav))
            for i, (fid, fname) in enumerate(nav):
                with crumb_cols[i]:
                    label = ("🏠 My Drive" if fid == "root" else f"📂 {fname}")
                    if st.button(label, key=f"crumb_{i}", use_container_width=True):
                        st.session_state["drive_nav"] = nav[:i + 1]
                        st.rerun()

            current_id   = nav[-1][0]
            current_name = nav[-1][1]

            # ── Subfolders grid ────────────────────────────────────────────────
            with st.spinner("Loading subfolders…"):
                subfolders = _list_drive_folders(creds, parent_id=current_id)

            if subfolders:
                st.markdown(f"**📂 Subfolders in *{current_name}*:**")
                cols = st.columns(3)
                for i, folder in enumerate(subfolders):
                    with cols[i % 3]:
                        if st.button(
                            f"📁 {folder['name']}",
                            key=f"sf_{folder['id']}",
                            use_container_width=True,
                        ):
                            st.session_state["drive_nav"].append(
                                (folder["id"], folder["name"])
                            )
                            st.rerun()
            else:
                st.caption(f"No subfolders in *{current_name}*.")

            # ── PDFs in current folder ─────────────────────────────────────────
            with st.spinner("Scanning for PDFs…"):
                pdf_files = _list_pdfs_in_folder(creds, current_id)

            st.markdown("---")
            if pdf_files:
                inv_list = [f for f in pdf_files if not is_credit_note(f["name"])]
                cn_list  = [f for f in pdf_files if is_credit_note(f["name"])]

                st.markdown(
                    f"📄 **{len(pdf_files)} PDFs** found in 📂 *{current_name}* "
                    f"— <span class='inv-tag'>{len(inv_list)} invoices</span>, "
                    f"<span class='cn-tag'>{len(cn_list)} credit notes</span>",
                    unsafe_allow_html=True,
                )

                with st.expander("Preview file list"):
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(
                            f"**<span class='inv-tag'>📑 Invoices ({len(inv_list)})</span>**",
                            unsafe_allow_html=True,
                        )
                        for f in inv_list:
                            kb = int(f.get("size", 0)) // 1024
                            st.markdown(
                                f"<div class='file-row'>📄 {f['name']} <small>({kb} KB)</small></div>",
                                unsafe_allow_html=True,
                            )
                    with c2:
                        st.markdown(
                            f"**<span class='cn-tag'>🔖 Credit Notes ({len(cn_list)})</span>**",
                            unsafe_allow_html=True,
                        )
                        for f in cn_list:
                            kb = int(f.get("size", 0)) // 1024
                            st.markdown(
                                f"<div class='file-row'>📄 {f['name']} <small>({kb} KB)</small></div>",
                                unsafe_allow_html=True,
                            )
f
                if st.button(
                    f"▶  Download & Process {len(pdf_files)} PDFs from '{current_name}'",
                    type="primary",
                    use_container_width=True,
                    key="run_drive",
                ):
                    with st.spinner(f"Downloading {len(pdf_files)} PDFs from Google Drive…"):
                        tmp_dir = _download_pdfs_to_temp(creds, pdf_files)

                    result = run_processing(tmp_dir)

                    if result:
                        show_results(*result)

                    shutil.rmtree(tmp_dir, ignore_errors=True)

            else:
                st.info(
                    f"📭 No PDFs found in **{current_name}**. "
                    "Navigate into a subfolder using the buttons above."
                )
