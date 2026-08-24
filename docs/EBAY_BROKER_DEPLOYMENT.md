# LUCAS eBay Broker

The production eBay connection flow must use one hosted LUCAS broker. Desktop users should never configure eBay developer keys.

## Flow

1. Desktop LUCAS opens `https://lucas.mikeyscards.com/ebay/connect`.
2. The broker redirects the user to official eBay OAuth.
3. eBay redirects back to the registered broker callback.
4. The broker stores the eBay refresh token server-side.
5. The broker redirects back to the running desktop app with a LUCAS connection token.
6. Desktop LUCAS stores only that connection token and asks the broker for short-lived eBay access tokens when listing cards.

## Required Environment

Set these on the hosted broker service:

```bash
EBAY_ENV=production
EBAY_CLIENT_ID=...
EBAY_CLIENT_SECRET=...
EBAY_RUNAME=...
EBAY_OAUTH_SCOPES="https://api.ebay.com/oauth/api_scope https://api.ebay.com/oauth/api_scope/sell.account https://api.ebay.com/oauth/api_scope/sell.inventory"
LUCAS_EBAY_BROKER_PUBLIC_URL=https://lucas.mikeyscards.com/ebay
LUCAS_EBAY_BROKER_STATE_SECRET=...
LUCAS_EBAY_BROKER_STORE_PATH=/data/ebay_broker_connections.json
HOST=0.0.0.0
PORT=8788
```

`EBAY_RUNAME` must be the redirect name/URL registered in the eBay developer app for the broker callback, not an individual desktop machine.

## Local Docker Run

```bash
docker build -f Dockerfile.ebay-broker -t lucas-ebay-broker .
docker run --rm -p 8788:8788 --env-file .env -e LUCAS_EBAY_BROKER_STATE_SECRET=change-me lucas-ebay-broker
```

Health check:

```bash
curl http://127.0.0.1:8788/health
```

## Desktop Configuration

Desktop LUCAS is broker-first by default. It uses:

```bash
LUCAS_EBAY_BROKER_URL=https://lucas.mikeyscards.com/ebay
```

Only use local desktop OAuth for development:

```bash
LUCAS_EBAY_USE_LOCAL_OAUTH=1
```

Production users should not set `LUCAS_EBAY_USE_LOCAL_OAUTH`.
