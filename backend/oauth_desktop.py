"""
OAuth Desktop Flow for Nova AI
Handles browser-based OAuth login via https://nova-ai.work/login
"""
from __future__ import annotations

import json
import os
import sys
import threading
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from typing import Optional, Dict, Any

# Login page URL
LOGIN_URL = "https://nova-ai.work/login"
CALLBACK_PORT = 8765
CALLBACK_PATH = "/auth-callback"


def _get_user_data_dir() -> Path:
    """Get the user data directory for storing credentials."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        if appdata:
            user_dir = Path(appdata) / "Nova AI"
        else:
            user_dir = Path(__file__).resolve().parent
    else:
        user_dir = Path.home() / ".nova-ai"
    
    try:
        user_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        user_dir = Path(__file__).resolve().parent
    
    return user_dir


def _get_user_file_path() -> Path:
    """Get the path to user_account.json."""
    return _get_user_data_dir() / "user_account.json"


def get_stored_user() -> Optional[Dict[str, Any]]:
    """
    Returns the cached user from user_account.json, or None if not logged in.
    
    Returns dict with keys:
      - uid: string (Firebase user ID) - REQUIRED
      - name: string (display name)
      - email: string
      - tier: string ("Free", "Standard", "Pro")
      - photo_url: string (avatar URL)
      - idToken: string (Firebase ID token for API calls)
      - refreshToken: string (for token refresh)
    """
    user_file = _get_user_file_path()
    
    if not user_file.exists():
        return None
    
    try:
        data = json.loads(user_file.read_text(encoding="utf-8"))
        if data.get("uid"):
            return data
        return None
    except Exception:
        return None


def save_user(user_data: Dict[str, Any]) -> bool:
    """Save user data to user_account.json."""
    user_file = _get_user_file_path()
    
    try:
        user_file.write_text(
            json.dumps(user_data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        return True
    except Exception as e:
        print(f"Failed to save user data: {e}")
        return False


def logout_user() -> bool:
    """Deletes user_account.json to log out."""
    user_file = _get_user_file_path()
    
    try:
        if user_file.exists():
            user_file.unlink()
        return True
    except Exception as e:
        print(f"Failed to logout: {e}")
        return False


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """HTTP request handler for OAuth callback."""
    
    user_data: Optional[Dict[str, Any]] = None
    
    def log_message(self, format, *args):
        """Suppress HTTP server logs."""
        pass
    
    def do_GET(self):
        """Handle GET request from OAuth callback."""
        parsed = urlparse(self.path)
        
        if parsed.path == CALLBACK_PATH:
            # Parse query parameters
            params = parse_qs(parsed.query)
            
            # Extract user data from callback
            user_data = {
                "uid": params.get("uid", [""])[0],
                "name": params.get("name", [""])[0],
                "email": params.get("email", [""])[0],
                "tier": params.get("tier", ["Free"])[0],
                "photo_url": params.get("photo_url", [""])[0],
                "handle": params.get("handle", [""])[0],
                "idToken": params.get("idToken", [""])[0],
                "refreshToken": params.get("refreshToken", [""])[0],
            }
            
            if user_data["uid"]:
                # Save user data
                save_user(user_data)
                OAuthCallbackHandler.user_data = user_data
                
                # Send success response
                self.send_response(200)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                
                success_html = """
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <title>로그인 성공</title>
                    <style>
                        body {
                            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                            display: flex;
                            justify-content: center;
                            align-items: center;
                            height: 100vh;
                            margin: 0;
                            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                            color: white;
                        }
                        .container {
                            text-align: center;
                            padding: 40px;
                            background: rgba(255,255,255,0.1);
                            border-radius: 20px;
                            backdrop-filter: blur(10px);
                        }
                        h1 { font-size: 2em; margin-bottom: 10px; }
                        p { font-size: 1.2em; opacity: 0.9; }
                        .checkmark {
                            font-size: 4em;
                            margin-bottom: 20px;
                        }
                    </style>
                </head>
                <body>
                    <div class="container">
                        <div class="checkmark">✓</div>
                        <h1>로그인 성공!</h1>
                        <p>이 창을 닫고 Nova AI로 돌아가세요.</p>
                    </div>
                    <script>setTimeout(() => window.close(), 2000);</script>
                </body>
                </html>
                """
                self.wfile.write(success_html.encode("utf-8"))
            else:
                # Send error response
                self.send_response(400)
                self.send_header("Content-type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(b"<h1>Login failed - no user ID</h1>")
        else:
            self.send_response(404)
            self.end_headers()


def start_oauth_flow(timeout: int = 300) -> Optional[Dict[str, Any]]:
    """
    Start the OAuth login flow.
    
    1. Deletes any existing user_account.json
    2. Starts local HTTP server on port 8765
    3. Opens browser to login URL
    4. Waits up to `timeout` seconds for callback
    5. Returns user data if successful, None otherwise
    """
    # Clear previous session
    logout_user()
    OAuthCallbackHandler.user_data = None
    
    # Start local HTTP server
    server = HTTPServer(("127.0.0.1", CALLBACK_PORT), OAuthCallbackHandler)
    server.timeout = timeout
    
    # Build login URL with redirect
    redirect_uri = f"http://localhost:{CALLBACK_PORT}{CALLBACK_PATH}"
    login_url = f"{LOGIN_URL}?redirect_uri={redirect_uri}"
    
    # Open browser
    try:
        webbrowser.open(login_url)
    except Exception as e:
        print(f"Failed to open browser: {e}")
        return None
    
    # Wait for callback
    server_thread = threading.Thread(target=lambda: server.handle_request())
    server_thread.start()
    server_thread.join(timeout=timeout)
    
    try:
        server.server_close()
    except Exception:
        pass
    
    # After login, sync tier from Firebase
    user_data = OAuthCallbackHandler.user_data
    if user_data and user_data.get("uid"):
        try:
            from backend.firebase_profile import refresh_user_profile_from_firebase
            profile = refresh_user_profile_from_firebase()
            if profile and profile.get("tier"):
                user_data["tier"] = profile["tier"]
                save_user(user_data)
        except Exception:
            pass
    
    return user_data


def is_logged_in() -> bool:
    """Check if a user is currently logged in."""
    user = get_stored_user()
    return user is not None and bool(user.get("uid"))
