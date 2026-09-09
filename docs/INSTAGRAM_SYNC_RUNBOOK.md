# LUCAS Instagram Inventory Sync Runbook

Last verified: 2026-09-09

This file is the source of truth for LUCAS Instagram Inventory Sync. Do not change these IDs unless a fresh Graph API test proves the replacement can read media and publishing quota.

## Known Good Setup

- Facebook app: LUCAS Inventory
- App ID: `1043156651570082`
- Facebook Page: Mikeys Cards
- Facebook Page ID: `1090039820868692`
- Instagram account username: `mikeys_cards_inventory`
- Instagram account ID that works for posting: `17841465322546974`
- Instagram account ID that caused `#10 Application does not have permission for this action`: `17841415167583312`

In `.env`, keep:

```env
LUCAS_INSTAGRAM_APP_ID=1043156651570082
LUCAS_INSTAGRAM_USER_ID=17841465322546974
LUCAS_INSTAGRAM_ACCESS_TOKEN=<Mikeys Cards PAGE access token>
```

Important: despite the env variable name, `LUCAS_INSTAGRAM_ACCESS_TOKEN` must be the Mikeys Cards Page access token, not the personal user token.

## Required Token Shape

The token must be a Page access token for Page ID `1090039820868692`, generated from a user token that can manage the Page.

The token must allow these Graph API calls:

- `GET /17841465322546974?fields=id,username,name`
- `GET /17841465322546974/media?fields=id,permalink&limit=1`
- `GET /17841465322546974/content_publishing_limit`

On 2026-09-09, the current Page token passed all three calls with Instagram account ID `17841465322546974`, and Meta `debug_token` reported `type=PAGE`, `is_valid=True`, `profile_id=1090039820868692`, and `expires_at` empty/unknown.

## Required Permissions

When generating the user token in Graph API Explorer, include the permissions needed for Page lookup and Instagram publishing:

- `pages_show_list`
- `business_management`
- `pages_read_engagement`
- `pages_read_user_content`
- `pages_manage_posts`
- `instagram_basic`
- `instagram_manage_comments`
- `instagram_content_publish`

The Page token result from:

```text
GET /me/accounts?fields=name,id,access_token,tasks
```

must return the Mikeys Cards Page with tasks including:

- `MANAGE`
- `CREATE_CONTENT`
- `MODERATE`
- `MESSAGING`
- `ADVERTISE`
- `ANALYZE`

## Exact Renewal Steps

1. Open Graph API Explorer:
   [https://developers.facebook.com/tools/explorer/](https://developers.facebook.com/tools/explorer/)

2. Select Meta App:
   `LUCAS Inventory`

3. Generate a User Access Token with the required permissions listed above.

4. If using a short-lived user token, exchange it for a longer-lived user token:

   ```text
   https://graph.facebook.com/v26.0/oauth/access_token?grant_type=fb_exchange_token&client_id=1043156651570082&client_secret=<APP_SECRET>&fb_exchange_token=<USER_ACCESS_TOKEN>
   ```

   Exchange token meaning: paste the user access token from Graph API Explorer into `fb_exchange_token`. This step turns a short-lived user token into a longer-lived user token. It is still not the final token LUCAS should use.

5. Put the longer-lived user token into Graph API Explorer's Access Token box.

6. Run:

   ```text
   GET /me/accounts?fields=name,id,access_token,tasks
   ```

7. Copy the `access_token` from the `Mikeys Cards` Page row. That Page token is the token LUCAS needs.

8. Update `.env`:

   ```env
   LUCAS_INSTAGRAM_APP_ID=1043156651570082
   LUCAS_INSTAGRAM_USER_ID=17841465322546974
   LUCAS_INSTAGRAM_ACCESS_TOKEN=<Mikeys Cards PAGE access token>
   ```

9. Validate before restarting LUCAS:

   ```bash
   python3 scripts/validate_instagram_env.py
   ```

   Do not call the renewal done unless the validator prints:

   ```text
   token type: PAGE
   token valid: True
   token profile/page id: 1090039820868692
   token expires: never/unknown
   profile: OK
   media: OK
   quota: OK
   page link: OK
   ```

10. Restart LUCAS / backend so `.env` reloads.

## Drift Check

If Instagram Inventory Sync suddenly fails after working:

1. First check `.env`.
2. If `LUCAS_INSTAGRAM_USER_ID` is `17841415167583312`, change it back to `17841465322546974`.
3. Run `python3 scripts/validate_instagram_env.py`.
4. Only regenerate tokens if the validator says the token is expired/invalid or media/quota access fails for the known-good ID.

Do not use a plain user token directly in `LUCAS_INSTAGRAM_ACCESS_TOKEN`. It may validate as a token but fail posting with Page/Instagram permission errors.

## 2026-09-09 Failure Note

The 2026-09-07 token in `.env` was the correct Page token shape, but it was still short-lived. Meta `debug_token` showed:

- `type=PAGE`
- `profile_id=1090039820868692`
- `is_valid=False`
- expired Monday, 2026-09-07 22:00 PDT

The backup file `.env.backup-before-page-token-update-20260907-final` held the longer-lived user token. Using that user token to run `/me/accounts?fields=name,id,access_token,tasks` produced a Page token that Meta reported as valid with no explicit expiry. That Page token replaced the expired `.env` token on 2026-09-09.
