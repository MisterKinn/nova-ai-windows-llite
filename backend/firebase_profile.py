"""
Firebase Profile Management for Nova AI
Handles Firestore REST API integration for user profiles and AI usage tracking.
"""
from __future__ import annotations

import json
import os
import time
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

from backend.oauth_desktop import get_stored_user, save_user, _get_user_data_dir


# Firebase Configuration
FIREBASE_CONFIG = {
    "apiKey": "AIzaSyDrxZaYCCy-jb8jmbCNVAjnoL6Ks866WLM",
    "projectId": "formulite-5b963",
}

# Plan/Tier Limits (daily AI calls)
PLAN_LIMITS = {
    'free': 5,
    'Free': 5,
    'standard': 50,
    'Standard': 50,
    'plus': 220,      # Alternative name for Standard
    'Plus': 220,
    'pro': 660,
    'Pro': 660,
}

# Cache for Firebase data (avoid repeated API calls)
_firebase_cache: Dict[str, Any] = {
    "profile": None,
    "usage": None,
    "last_refresh": 0,
}
_CACHE_TTL = 60  # seconds


class FirebaseProfileError(Exception):
    """Firebase profile operation error."""
    pass


def _get_local_usage_path() -> Path:
    """Get path to local usage tracking file."""
    return _get_user_data_dir() / "local_usage.json"


def _get_local_usage() -> Dict[str, Any]:
    """Get local usage data (fallback when Firebase unavailable)."""
    path = _get_local_usage_path()
    if not path.exists():
        return {"date": "", "usage": 0}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"date": "", "usage": 0}


def _save_local_usage(data: Dict[str, Any]) -> None:
    """Save local usage data."""
    path = _get_local_usage_path()
    try:
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def get_valid_id_token() -> Optional[str]:
    """
    Get a valid Firebase ID token, refreshing if expired.
    Used for authenticated API calls.
    """
    user = get_stored_user()
    if not user:
        return None
    
    id_token = user.get("idToken")
    refresh_token = user.get("refreshToken")
    
    if not id_token:
        return None
    
    # Try to decode token to check expiry (simple check)
    # In production, you'd properly decode the JWT
    # For now, we'll try to refresh if we have a refresh token
    
    return id_token


def refresh_id_token() -> Optional[str]:
    """Refresh the Firebase ID token using the refresh token."""
    if not REQUESTS_AVAILABLE:
        return None
    
    user = get_stored_user()
    if not user or not user.get("refreshToken"):
        return None
    
    try:
        url = f"https://securetoken.googleapis.com/v1/token?key={FIREBASE_CONFIG['apiKey']}"
        payload = {
            "grant_type": "refresh_token",
            "refresh_token": user["refreshToken"]
        }
        
        response = requests.post(url, data=payload, timeout=10)
        data = response.json()
        
        if "id_token" in data:
            # Update stored user with new tokens
            user["idToken"] = data["id_token"]
            if "refresh_token" in data:
                user["refreshToken"] = data["refresh_token"]
            save_user(user)
            return data["id_token"]
        
        return None
    except Exception as e:
        print(f"Token refresh failed: {e}")
        return None


def refresh_user_profile_from_firebase() -> Optional[Dict[str, Any]]:
    """
    Called at app startup to sync latest user data from Firestore.
    
    1. Reads uid from user_account.json
    2. Fetches users/{uid} document from Firestore REST API
    3. Updates local tier/displayName if changed
    4. Returns updated profile dict
    
    Returns dict with keys:
      - uid, tier, display_name, aiCallUsage, email, photo_url
    """
    if not REQUESTS_AVAILABLE:
        return None
    
    user = get_stored_user()
    if not user or not user.get("uid"):
        return None
    
    uid = user["uid"]
    id_token = get_valid_id_token()
    
    if not id_token:
        # Try to refresh token
        id_token = refresh_id_token()
        if not id_token:
            return None
    
    try:
        project_id = FIREBASE_CONFIG["projectId"]
        url = f"https://firestore.googleapis.com/v1/projects/{project_id}/databases/(default)/documents/users/{uid}"
        headers = {"Authorization": f"Bearer {id_token}"}
        
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            doc = response.json()
            fields = doc.get("fields", {})
            
            # Parse Firestore field values
            def get_value(field: Dict) -> Any:
                for key in ["stringValue", "integerValue", "booleanValue", "doubleValue"]:
                    if key in field:
                        val = field[key]
                        if key == "integerValue":
                            return int(val)
                        return val
                return None
            
            profile = {
                "uid": uid,
                "tier": get_value(fields.get("tier", {})) or user.get("tier", "Free"),
                "display_name": get_value(fields.get("displayName", {})) or user.get("name", ""),
                "email": get_value(fields.get("email", {})) or user.get("email", ""),
                "photo_url": get_value(fields.get("photoURL", {})) or user.get("photo_url", ""),
                "aiCallUsage": get_value(fields.get("aiCallUsage", {})) or 0,
            }
            
            # Update local cache if tier changed
            if profile["tier"] != user.get("tier"):
                user["tier"] = profile["tier"]
                save_user(user)
            
            return profile
        
        elif response.status_code == 401:
            # Token expired, try to refresh
            id_token = refresh_id_token()
            if id_token:
                return refresh_user_profile_from_firebase()  # Retry once
            return None
        
        else:
            print(f"Firebase profile fetch failed: {response.status_code}")
            return None
            
    except Exception as e:
        print(f"Firebase profile refresh error: {e}")
        return None


def get_user_profile(uid: str) -> Optional[Dict[str, Any]]:
    """
    Get full user profile from Firestore.
    Returns: display_name, email, photo_url, uid, tier, aiCallUsage
    """
    if not uid:
        return None
    
    # Try Firebase first
    profile = refresh_user_profile_from_firebase()
    if profile:
        return profile
    
    # Fallback to cached data
    user = get_stored_user()
    if user and user.get("uid") == uid:
        return {
            "uid": uid,
            "tier": user.get("tier", "Free"),
            "display_name": user.get("name", ""),
            "email": user.get("email", ""),
            "photo_url": user.get("photo_url", ""),
            "aiCallUsage": 0,
        }
    
    return None


def get_ai_usage(uid: str) -> int:
    """
    Get current AI call usage count for the user.
    Uses cached data if available and fresh, otherwise fetches from Firebase.
    """
    if not uid:
        return 0
    
    now = time.time()
    
    # Return cached usage if fresh
    if _firebase_cache["usage"] is not None and (now - _firebase_cache["last_refresh"]) < _CACHE_TTL:
        return _firebase_cache["usage"]
    
    # Try to get from Firebase
    profile = refresh_user_profile_from_firebase()
    if profile:
        usage = profile.get("aiCallUsage", 0)
        _firebase_cache["usage"] = usage
        _firebase_cache["last_refresh"] = now
        return usage
    
    # Fallback to local tracking
    local = _get_local_usage()
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Reset if new day
    if local.get("date") != today:
        local = {"date": today, "usage": 0}
        _save_local_usage(local)
    
    return local.get("usage", 0)


def force_refresh_usage() -> int:
    """Force refresh usage from Firebase, bypassing cache."""
    _firebase_cache["usage"] = None
    _firebase_cache["last_refresh"] = 0
    user = get_stored_user()
    if user and user.get("uid"):
        return get_ai_usage(user["uid"])
    return 0


def increment_ai_usage(uid: str) -> bool:
    """
    Atomically increment aiCallUsage field in Firestore.
    Called each time user makes an AI request.
    Returns True if successful.
    """
    if not uid:
        return False
    
    if not REQUESTS_AVAILABLE:
        # Local fallback
        return _increment_local_usage()
    
    id_token = get_valid_id_token()
    if not id_token:
        id_token = refresh_id_token()
    
    if not id_token:
        # Fallback to local tracking
        return _increment_local_usage()
    
    try:
        project_id = FIREBASE_CONFIG["projectId"]
        
        # First, get current usage
        url = f"https://firestore.googleapis.com/v1/projects/{project_id}/databases/(default)/documents/users/{uid}"
        headers = {
            "Authorization": f"Bearer {id_token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            doc = response.json()
            fields = doc.get("fields", {})
            current_usage = 0
            
            if "aiCallUsage" in fields:
                val = fields["aiCallUsage"].get("integerValue", 0)
                current_usage = int(val)
            
            # Update with incremented value
            new_usage = current_usage + 1
            update_url = f"{url}?updateMask.fieldPaths=aiCallUsage"
            payload = {
                "fields": {
                    "aiCallUsage": {"integerValue": str(new_usage)}
                }
            }
            
            update_response = requests.patch(update_url, headers=headers, json=payload, timeout=10)
            
            if update_response.status_code in (200, 201):
                # Update cache
                _firebase_cache["usage"] = new_usage
                _firebase_cache["last_refresh"] = time.time()
                return True
            else:
                print(f"Usage update failed: {update_response.status_code}")
                return _increment_local_usage()
        
        elif response.status_code == 404:
            # Document doesn't exist, create it
            create_url = f"https://firestore.googleapis.com/v1/projects/{project_id}/databases/(default)/documents/users?documentId={uid}"
            payload = {
                "fields": {
                    "aiCallUsage": {"integerValue": "1"}
                }
            }
            create_response = requests.post(create_url, headers=headers, json=payload, timeout=10)
            return create_response.status_code in (200, 201)
        
        else:
            return _increment_local_usage()
            
    except Exception as e:
        print(f"Firebase usage increment error: {e}")
        return _increment_local_usage()


def _increment_local_usage() -> bool:
    """Increment usage in local tracking file (fallback)."""
    local = _get_local_usage()
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Reset if new day
    if local.get("date") != today:
        local = {"date": today, "usage": 0}
    
    local["usage"] = local.get("usage", 0) + 1
    _save_local_usage(local)
    return True


def get_remaining_usage(uid: str, tier: str = "Free") -> int:
    """
    Get remaining AI calls for today.
    """
    current_usage = get_ai_usage(uid)
    limit = PLAN_LIMITS.get(tier, PLAN_LIMITS.get("Free", 5))
    return max(0, limit - current_usage)


def check_usage_limit(uid: str, tier: str = "Free") -> bool:
    """
    Check if user has remaining AI calls.
    Returns True if user can make more calls.
    """
    return get_remaining_usage(uid, tier) > 0


def get_plan_limit(tier: str) -> int:
    """Get the daily limit for a given plan/tier."""
    return PLAN_LIMITS.get(tier, PLAN_LIMITS.get("Free", 5))
