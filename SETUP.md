{
  "mcpServers": {
    "ms365-mailer": {
      "command": "python3",
      "args": ["${CLAUDE_PLUGIN_ROOT}/servers/graph_mailer.py"],
      "env": {
        "GRAPH_TENANT_ID":     "PASTE_YOUR_DIRECTORY_TENANT_ID_HERE",
        "GRAPH_CLIENT_ID":     "PASTE_YOUR_APPLICATION_CLIENT_ID_HERE",
        "GRAPH_CLIENT_SECRET": "PASTE_YOUR_CLIENT_SECRET_VALUE_HERE",
        "GRAPH_SEND_AS":        "andrew.hurnard@thecachegroup.com.au",
        "GRAPH_SENDER_NAME":    "Andrew Hurnard",
        "GRAPH_SENDER_TITLE":   "Director",
        "GRAPH_SENDER_COMPANY": "The Cache Group",
        "GRAPH_SENDER_PHONE":   "0417 037 451",
        "GRAPH_SIGN_OFF":       "Andrew"
      }
    }
  }
}
