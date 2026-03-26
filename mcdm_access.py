from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, List, Mapping, Sequence

import streamlit as st

try:
    from sqlalchemy import create_engine, text
    SQLALCHEMY_AVAILABLE = True
except Exception:
    SQLALCHEMY_AVAILABLE = False
    create_engine = None
    text = None


SESSION_TOKEN_KEY = "_mcdm_access_session_token"
USER_BOOTSTRAPPED_KEY = "_mcdm_access_user_bootstrapped"
EVENT_CACHE_KEY = "_mcdm_access_event_cache"
TRACKING_ERROR_KEY = "_mcdm_access_tracking_error"


@dataclass(frozen=True)
class AuthSettings:
    require_login: bool
    enabled: bool
    configured: bool
    provider: str | None
    signup_provider: str | None
    admin_emails: frozenset[str]
    privacy_notice_tr: str
    privacy_notice_en: str


@dataclass(frozen=True)
class AnalyticsSettings:
    enabled: bool
    db_url: str | None
    table_prefix: str


@dataclass(frozen=True)
class CurrentUser:
    is_logged_in: bool
    email: str
    name: str
    auth_subject: str
    email_verified: bool
    raw: Mapping[str, Any]


@dataclass(frozen=True)
class UserRecord:
    display_name: str
    email_verified: bool
    login_count: int
    email_grace_used: bool


def preserved_session_keys() -> set[str]:
    return {
        "ui_lang",
        "show_step_guidance",
        SESSION_TOKEN_KEY,
        USER_BOOTSTRAPPED_KEY,
        EVENT_CACHE_KEY,
        TRACKING_ERROR_KEY,
        "_mcdm_last_activity_at",
        "_mcdm_email_grace_given",
        "_mcdm_name_collected",
        "_mcdm_login_count",
    }


def _section(name: str) -> Dict[str, Any]:
    try:
        value = st.secrets[name]
    except Exception:
        return {}
    if hasattr(value, "to_dict"):
        return value.to_dict()
    try:
        return dict(value)
    except Exception:
        return {}


def _nested_section(parent: str, child: str | None) -> Dict[str, Any]:
    if not child:
        return {}
    base = _section(parent)
    value = base.get(child)
    if hasattr(value, "to_dict"):
        return value.to_dict()
    if isinstance(value, dict):
        return value
    try:
        return dict(value)
    except Exception:
        return {}


def _as_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "on"}
    return bool(value)


def _as_list(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [item.strip() for item in value.split(",") if item.strip()]
    if isinstance(value, Iterable):
        return [str(item).strip() for item in value if str(item).strip()]
    return []


def _streamlit_auth_supported() -> bool:
    return hasattr(st, "user") and callable(getattr(st, "login", None)) and callable(getattr(st, "logout", None))


def get_auth_settings() -> AuthSettings:
    cfg = _section("mcdm_auth")
    provider = str(cfg.get("provider") or "").strip() or None
    signup_provider = str(cfg.get("signup_provider") or "").strip() or provider
    require_login = _as_bool(cfg.get("require_login"), default=False)
    admin_emails = frozenset(email.lower() for email in _as_list(cfg.get("admin_emails")))
    auth_root = _section("auth")
    has_shared = bool(auth_root.get("redirect_uri")) and bool(auth_root.get("cookie_secret"))
    if provider:
        provider_cfg = _nested_section("auth", provider)
        configured = has_shared and bool(provider_cfg.get("client_id")) and bool(provider_cfg.get("client_secret")) and bool(provider_cfg.get("server_metadata_url"))
    else:
        configured = has_shared and bool(auth_root.get("client_id")) and bool(auth_root.get("client_secret")) and bool(auth_root.get("server_metadata_url"))
    enabled = configured and _streamlit_auth_supported()
    return AuthSettings(
        require_login=require_login,
        enabled=enabled,
        configured=configured,
        provider=provider,
        signup_provider=signup_provider,
        admin_emails=admin_emails,
        privacy_notice_tr=str(
            cfg.get("privacy_notice_tr")
            or "Bu uygulama giris yapan kullanicilarin e-posta adresini ve temel kullanim olaylarini kaydeder."
        ),
        privacy_notice_en=str(
            cfg.get("privacy_notice_en")
            or "This app stores signed-in users' email addresses and basic usage events."
        ),
    )


def get_analytics_settings() -> AnalyticsSettings:
    cfg = _section("mcdm_analytics")
    try:
        _secret_db = str(st.secrets.get("database_url", "") or "").strip()
    except Exception:
        _secret_db = ""
    db_url = (
        str(cfg.get("db_url") or "").strip()
        or _secret_db
        or str(os.environ.get("MCDM_ANALYTICS_DB_URL", "")).strip()
        or None
    )
    enabled = _as_bool(cfg.get("enabled"), default=True) and bool(db_url) and SQLALCHEMY_AVAILABLE
    table_prefix = str(cfg.get("table_prefix") or "mcdm").strip() or "mcdm"
    return AnalyticsSettings(enabled=enabled, db_url=db_url, table_prefix=table_prefix)


def get_current_user() -> CurrentUser:
    user = getattr(st, "user", None)
    is_logged_in = bool(getattr(user, "is_logged_in", False))
    if not is_logged_in:
        return CurrentUser(
            is_logged_in=False,
            email="",
            name="",
            auth_subject="",
            email_verified=False,
            raw={},
        )
    raw = {}
    try:
        raw = dict(user)
    except Exception:
        pass
    email = str(getattr(user, "email", "") or raw.get("email") or "").strip().lower()
    name = str(getattr(user, "name", "") or raw.get("name") or email).strip()
    auth_subject = str(getattr(user, "sub", "") or raw.get("sub") or email).strip()
    email_verified = bool(getattr(user, "email_verified", raw.get("email_verified", False)))
    stored = get_user_record(auth_subject, email)
    if stored:
        stored_name = str(stored.display_name or "").strip()
        if stored_name and (not name or name.lower() == email.lower() or "@" in name):
            name = stored_name
        email_verified = bool(email_verified or stored.email_verified)
    return CurrentUser(
        is_logged_in=True,
        email=email,
        name=name,
        auth_subject=auth_subject,
        email_verified=email_verified,
        raw=raw,
    )


@st.cache_resource(show_spinner=False)
def _get_engine(db_url: str):
    if not SQLALCHEMY_AVAILABLE:
        raise RuntimeError("SQLAlchemy is not available.")
    return create_engine(db_url, pool_pre_ping=True, future=True)


def _table_names(settings: AnalyticsSettings) -> Dict[str, str]:
    prefix = settings.table_prefix
    return {
        "users": f"{prefix}_users",
        "sessions": f"{prefix}_sessions",
        "events": f"{prefix}_events",
    }


def _ensure_schema(settings: AnalyticsSettings) -> None:
    if not settings.enabled or not settings.db_url:
        return
    tables = _table_names(settings)
    engine = _get_engine(settings.db_url)
    ddl = [
        f"""
        CREATE TABLE IF NOT EXISTS {tables["users"]} (
            auth_subject TEXT PRIMARY KEY,
            email TEXT NOT NULL,
            display_name TEXT,
            email_verified BOOLEAN NOT NULL DEFAULT FALSE,
            email_grace_used BOOLEAN NOT NULL DEFAULT FALSE,
            created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            last_login_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            login_count INTEGER NOT NULL DEFAULT 0
        )
        """,
        f"ALTER TABLE {tables['users']} ADD COLUMN IF NOT EXISTS email_grace_used BOOLEAN NOT NULL DEFAULT FALSE",
        f"CREATE INDEX IF NOT EXISTS {tables['users']}_email_idx ON {tables['users']} (LOWER(email))",
        f"""
        CREATE TABLE IF NOT EXISTS {tables["sessions"]} (
            session_token TEXT PRIMARY KEY,
            auth_subject TEXT NOT NULL,
            user_email TEXT NOT NULL,
            display_name TEXT,
            started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            last_seen_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            app_version TEXT
        )
        """,
        f"CREATE INDEX IF NOT EXISTS {tables['sessions']}_email_idx ON {tables['sessions']} (LOWER(user_email))",
        f"""
        CREATE TABLE IF NOT EXISTS {tables["events"]} (
            event_id TEXT PRIMARY KEY,
            session_token TEXT NOT NULL,
            auth_subject TEXT,
            user_email TEXT,
            event_type TEXT NOT NULL,
            occurred_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
            payload_json JSONB
        )
        """,
        f"CREATE INDEX IF NOT EXISTS {tables['events']}_event_type_idx ON {tables['events']} (event_type)",
        f"CREATE INDEX IF NOT EXISTS {tables['events']}_occurred_at_idx ON {tables['events']} (occurred_at DESC)",
        f"CREATE INDEX IF NOT EXISTS {tables['events']}_user_email_idx ON {tables['events']} (LOWER(user_email))",
    ]
    with engine.begin() as conn:
        for statement in ddl:
            conn.exec_driver_sql(statement)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _remember_tracking_error(exc: Exception) -> None:
    if TRACKING_ERROR_KEY not in st.session_state:
        st.session_state[TRACKING_ERROR_KEY] = str(exc)


def tracking_error() -> str | None:
    value = st.session_state.get(TRACKING_ERROR_KEY)
    return str(value).strip() if value else None


def _get_session_token() -> str:
    token = str(st.session_state.get(SESSION_TOKEN_KEY) or "").strip()
    if token:
        return token
    token = uuid.uuid4().hex
    st.session_state[SESSION_TOKEN_KEY] = token
    return token


def is_admin_email(email: str, settings: AuthSettings | None = None) -> bool:
    auth_settings = settings or get_auth_settings()
    return bool(email and email.lower() in auth_settings.admin_emails)


def get_user_record(
    auth_subject: str,
    email: str = "",
    analytics_settings: AnalyticsSettings | None = None,
) -> UserRecord | None:
    if not auth_subject and not email:
        return None
    settings = analytics_settings or get_analytics_settings()
    if not settings.enabled or not settings.db_url:
        return None
    tables = _table_names(settings)
    try:
        _ensure_schema(settings)
        engine = _get_engine(settings.db_url)
        with engine.begin() as conn:
            row = None
            if auth_subject:
                row = conn.execute(
                    text(
                        f"""
                        SELECT
                            COALESCE(display_name, '') AS display_name,
                            email_verified,
                            login_count,
                            COALESCE(email_grace_used, FALSE) AS email_grace_used
                        FROM {tables["users"]}
                        WHERE auth_subject = :auth_subject
                        """
                    ),
                    {"auth_subject": auth_subject},
                ).fetchone()
            if row is None and email:
                row = conn.execute(
                    text(
                        f"""
                        SELECT
                            COALESCE(display_name, '') AS display_name,
                            email_verified,
                            login_count,
                            COALESCE(email_grace_used, FALSE) AS email_grace_used
                        FROM {tables["users"]}
                        WHERE LOWER(email) = LOWER(:email)
                        ORDER BY last_login_at DESC
                        LIMIT 1
                        """
                    ),
                    {"email": email},
                ).fetchone()
    except Exception as exc:
        _remember_tracking_error(exc)
        return None
    if row is None:
        return None
    payload = dict(row._mapping)
    return UserRecord(
        display_name=str(payload.get("display_name") or "").strip(),
        email_verified=bool(payload.get("email_verified", False)),
        login_count=int(payload.get("login_count") or 0),
        email_grace_used=bool(payload.get("email_grace_used", False)),
    )


def bootstrap_user_session(user: CurrentUser, analytics_settings: AnalyticsSettings | None = None) -> int:
    """Upserts the user into the DB, increments login_count, and returns login_count.

    Returns the login_count from the database (1 on first login, >1 on subsequent).
    Falls back to the cached value in session_state if already bootstrapped.
    """
    if not user.is_logged_in:
        return 1
    cached = st.session_state.get("_mcdm_login_count")
    if st.session_state.get(USER_BOOTSTRAPPED_KEY) and cached is not None:
        return int(cached)
    settings = analytics_settings or get_analytics_settings()
    session_token = _get_session_token()
    login_count = 1
    if settings.enabled and settings.db_url:
        tables = _table_names(settings)
        try:
            _ensure_schema(settings)
            engine = _get_engine(settings.db_url)
            now_ts = _now_iso()
            with engine.begin() as conn:
                row = conn.execute(
                    text(
                        f"""
                        INSERT INTO {tables["users"]} (
                            auth_subject, email, display_name, email_verified, email_grace_used, created_at, last_login_at, login_count
                        )
                        VALUES (
                            :auth_subject, :email, :display_name, :email_verified, FALSE, :created_at, :last_login_at, 1
                        )
                        ON CONFLICT (auth_subject) DO UPDATE SET
                            email = EXCLUDED.email,
                            display_name = CASE
                                WHEN COALESCE(EXCLUDED.display_name, '') = '' OR POSITION('@' IN EXCLUDED.display_name) > 0
                                    THEN {tables["users"]}.display_name
                                ELSE EXCLUDED.display_name
                            END,
                            email_verified = EXCLUDED.email_verified,
                            last_login_at = EXCLUDED.last_login_at,
                            login_count = {tables["users"]}.login_count + 1
                        RETURNING login_count
                        """
                    ),
                    {
                        "auth_subject": user.auth_subject,
                        "email": user.email,
                        "display_name": user.name,
                        "email_verified": user.email_verified,
                        "created_at": now_ts,
                        "last_login_at": now_ts,
                    },
                ).fetchone()
                if row is not None:
                    login_count = int(row[0])
                conn.execute(
                    text(
                        f"""
                        INSERT INTO {tables["sessions"]} (
                            session_token, auth_subject, user_email, display_name, started_at, last_seen_at, app_version
                        )
                        VALUES (
                            :session_token, :auth_subject, :user_email, :display_name, :started_at, :last_seen_at, :app_version
                        )
                        ON CONFLICT (session_token) DO UPDATE SET
                            auth_subject = EXCLUDED.auth_subject,
                            user_email = EXCLUDED.user_email,
                            display_name = EXCLUDED.display_name,
                            last_seen_at = EXCLUDED.last_seen_at,
                            app_version = EXCLUDED.app_version
                        """
                    ),
                    {
                        "session_token": session_token,
                        "auth_subject": user.auth_subject,
                        "user_email": user.email,
                        "display_name": user.name,
                        "started_at": now_ts,
                        "last_seen_at": now_ts,
                        "app_version": str(os.environ.get("MCDM_APP_VERSION", "streamlit")),
                    },
                )
        except Exception as exc:
            _remember_tracking_error(exc)
    st.session_state["_mcdm_login_count"] = login_count
    st.session_state[USER_BOOTSTRAPPED_KEY] = True
    track_event("session_started", {"email_verified": user.email_verified, "login_count": login_count}, settings=settings, user=user)
    return login_count


def consume_email_verification_grace(
    auth_subject: str,
    analytics_settings: AnalyticsSettings | None = None,
) -> bool:
    if not auth_subject:
        return False
    settings = analytics_settings or get_analytics_settings()
    if not settings.enabled or not settings.db_url:
        if st.session_state.get("_mcdm_email_grace_given"):
            return False
        st.session_state["_mcdm_email_grace_given"] = True
        return True
    tables = _table_names(settings)
    try:
        _ensure_schema(settings)
        engine = _get_engine(settings.db_url)
        with engine.begin() as conn:
            row = conn.execute(
                text(
                    f"""
                    UPDATE {tables["users"]}
                    SET email_grace_used = TRUE
                    WHERE auth_subject = :auth_subject
                      AND COALESCE(email_grace_used, FALSE) = FALSE
                    RETURNING auth_subject
                    """
                ),
                {"auth_subject": auth_subject},
            ).fetchone()
    except Exception as exc:
        _remember_tracking_error(exc)
        if st.session_state.get("_mcdm_email_grace_given"):
            return False
        st.session_state["_mcdm_email_grace_given"] = True
        return True
    return row is not None


def track_event(
    event_type: str,
    payload: Mapping[str, Any] | None = None,
    *,
    once_key: str | None = None,
    settings: AnalyticsSettings | None = None,
    user: CurrentUser | None = None,
) -> None:
    cache = set(st.session_state.get(EVENT_CACHE_KEY) or set())
    if once_key and once_key in cache:
        return
    analytics_settings = settings or get_analytics_settings()
    current_user = user or get_current_user()
    if not current_user.is_logged_in:
        return
    if once_key:
        cache.add(once_key)
        st.session_state[EVENT_CACHE_KEY] = cache
    if not analytics_settings.enabled or not analytics_settings.db_url:
        return
    tables = _table_names(analytics_settings)
    session_token = _get_session_token()
    event_payload = dict(payload or {})
    event_payload.setdefault("logged_at", _now_iso())
    try:
        _ensure_schema(analytics_settings)
        engine = _get_engine(analytics_settings.db_url)
        with engine.begin() as conn:
            conn.execute(
                text(
                    f"""
                    UPDATE {tables["sessions"]}
                    SET last_seen_at = :last_seen_at
                    WHERE session_token = :session_token
                    """
                ),
                {
                    "last_seen_at": _now_iso(),
                    "session_token": session_token,
                },
            )
            conn.execute(
                text(
                    f"""
                    INSERT INTO {tables["events"]} (
                        event_id, session_token, auth_subject, user_email, event_type, occurred_at, payload_json
                    )
                    VALUES (
                        :event_id, :session_token, :auth_subject, :user_email, :event_type, :occurred_at, CAST(:payload_json AS JSONB)
                    )
                    """
                ),
                {
                    "event_id": uuid.uuid4().hex,
                    "session_token": session_token,
                    "auth_subject": current_user.auth_subject,
                    "user_email": current_user.email,
                    "event_type": event_type,
                    "occurred_at": _now_iso(),
                    "payload_json": json.dumps(event_payload, ensure_ascii=True, default=str),
                },
            )
    except Exception as exc:
        _remember_tracking_error(exc)


def logout_user() -> None:
    for key in (
        SESSION_TOKEN_KEY,
        USER_BOOTSTRAPPED_KEY,
        EVENT_CACHE_KEY,
        "_mcdm_last_activity_at",
        "_mcdm_email_grace_given",
        "_mcdm_name_collected",
        "_mcdm_login_count",
    ):
        st.session_state.pop(key, None)
    if callable(getattr(st, "logout", None)):
        st.logout()


def update_display_name(auth_subject: str, name: str, analytics_settings: AnalyticsSettings | None = None) -> None:
    """Updates the display_name for a user in the database."""
    if not auth_subject or not name:
        return
    settings = analytics_settings or get_analytics_settings()
    if not settings.enabled or not settings.db_url:
        return
    tables = _table_names(settings)
    try:
        engine = _get_engine(settings.db_url)
        with engine.begin() as conn:
            conn.execute(
                text(
                    f"UPDATE {tables['users']} SET display_name = :name WHERE auth_subject = :auth_subject"
                ),
                {"name": name.strip(), "auth_subject": auth_subject},
            )
    except Exception as exc:
        _remember_tracking_error(exc)


def login_user(provider: str | None = None) -> None:
    if provider:
        st.login(provider)
        return
    st.login()


def auth_gate_context() -> Dict[str, Any]:
    auth_settings = get_auth_settings()
    user = get_current_user()
    analytics_settings = get_analytics_settings()
    return {
        "auth_settings": auth_settings,
        "analytics_settings": analytics_settings,
        "user": user,
    }


def get_admin_summary(limit: int = 20) -> Dict[str, Any] | None:
    settings = get_analytics_settings()
    if not settings.enabled or not settings.db_url:
        return None
    tables = _table_names(settings)
    try:
        _ensure_schema(settings)
        engine = _get_engine(settings.db_url)
        with engine.begin() as conn:
            total_users = conn.execute(text(f"SELECT COUNT(DISTINCT LOWER(email)) FROM {tables['users']}")).scalar() or 0
            total_sessions = conn.execute(text(f"SELECT COUNT(*) FROM {tables['sessions']}")).scalar() or 0
            total_analyses = conn.execute(
                text(f"SELECT COUNT(*) FROM {tables['events']} WHERE event_type = 'analysis_completed'")
            ).scalar() or 0
            last_7d_users = conn.execute(
                text(
                    f"""
                    SELECT COUNT(DISTINCT LOWER(user_email))
                    FROM {tables["events"]}
                    WHERE occurred_at >= NOW() - INTERVAL '7 days'
                    """
                )
            ).scalar() or 0
            recent_rows = conn.execute(
                text(
                    f"""
                    SELECT
                        user_email,
                        MAX(occurred_at) AS last_seen_at,
                        COUNT(*) FILTER (WHERE event_type = 'analysis_completed') AS analyses_completed,
                        COUNT(*) FILTER (WHERE event_type = 'session_started') AS session_count
                    FROM {tables["events"]}
                    GROUP BY user_email
                    ORDER BY last_seen_at DESC
                    LIMIT :limit
                    """
                ),
                {"limit": int(max(1, limit))},
            )
            recent_users = [dict(row._mapping) for row in recent_rows]
            event_rows = conn.execute(
                text(
                    f"""
                    SELECT event_type, COUNT(*) AS total
                    FROM {tables["events"]}
                    GROUP BY event_type
                    ORDER BY total DESC, event_type ASC
                    """
                )
            )
            events_by_type = [dict(row._mapping) for row in event_rows]
    except Exception as exc:
        _remember_tracking_error(exc)
        return None
    return {
        "total_users": int(total_users),
        "total_sessions": int(total_sessions),
        "total_analyses": int(total_analyses),
        "last_7d_users": int(last_7d_users),
        "recent_users": recent_users,
        "events_by_type": events_by_type,
    }
