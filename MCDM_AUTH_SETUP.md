# MCDM Toolbox Auth and Analytics Setup

This app now supports:

- mandatory sign-in before analysis
- email/password sign-up through Auth0
- user email collection
- usage analytics stored in Postgres
- admin-only usage summary inside the sidebar

## Recommended stack

- Auth: Auth0 Universal Login
- Analytics DB: Supabase Postgres
- App host: Streamlit Community Cloud

## 1. Auth0 configuration

Create an Auth0 application for a regular web app.

Allowed callback URLs:

- `https://mcdm-assistance.streamlit.app/oauth2callback`
- `http://localhost:8501/oauth2callback`

Allowed logout URLs:

- `https://mcdm-assistance.streamlit.app`
- `http://localhost:8501`

Enable a Database Connection in Auth0 so users can sign up with email and password.

Optional but recommended:

- enable email verification
- customize Universal Login branding

## 2. Supabase configuration

Create a Supabase project and copy the Postgres connection string.

Use the pooled or direct Postgres connection string in Streamlit secrets as `mcdm_analytics.db_url`.

The app will auto-create these tables on first successful connection:

- `mcdm_users`
- `mcdm_sessions`
- `mcdm_events`

## 3. Streamlit secrets

Use `.streamlit/secrets.toml` based on `.streamlit/secrets.template.toml`.

Important:

- `mcdm_auth.require_login = true` enables the login wall
- `mcdm_auth.provider = "auth0"` is the normal login provider
- `mcdm_auth.signup_provider = "auth0signup"` opens Auth0 on the sign-up screen

## 4. Auth0 provider naming in Streamlit

The app expects two provider blocks:

- `[auth.auth0]` for standard login
- `[auth.auth0signup]` for sign-up-first flow

The `auth0signup` block should reuse the same Auth0 app but set:

- `client_kwargs = { "screen_hint" = "signup" }`

## 5. Admin access

Add your email under:

- `mcdm_auth.admin_emails = ["you@example.com"]`

Admins can see the usage summary panel in the sidebar.

## 6. Deployment sequence

1. Update `requirements.txt`.
2. Add Streamlit secrets in Community Cloud.
3. Redeploy the app.
4. Test sign-up.
5. Test sign-in.
6. Run one sample analysis.
7. Check the admin usage panel.

## Notes

- Resetting an analysis does not log the user out.
- Logging failures do not block analysis.
- Logging out is handled only by the dedicated logout button.
