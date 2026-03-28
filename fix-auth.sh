#!/bin/bash
# Post-configure: patch auth mode to none in the generated config
STATE_DIR="${OPENCLAW_STATE_DIR:-/data/.openclaw}"
CONFIG="$STATE_DIR/openclaw.json"
if [ -f "$CONFIG" ]; then
  node -e "
    const fs = require('fs');
    const c = JSON.parse(fs.readFileSync('$CONFIG','utf8'));
    c.gateway = c.gateway || {};
    c.gateway.auth = { mode: 'none' };
    fs.writeFileSync('$CONFIG', JSON.stringify(c, null, 2));
    console.log('[fix-auth] set gateway.auth.mode = none');
  "
fi
