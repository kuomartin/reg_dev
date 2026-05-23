#!/bin/bash

# multi-push.sh
# Usage: ./multi-push.sh scriptId1 [scriptId2 ...]

# Check if scriptIds are provided
if [ "$#" -eq 0 ]; then
    echo "Usage: $0 scriptId1 [scriptId2 ...]"
    exit 1
fi

# Use absolute path to ensure it works from anywhere, or relative if preferred.
# Here we assume it's in the current directory as per the project structure.
CLASP_JSON="$(pwd)/.clasp.json"

if [ ! -f "$CLASP_JSON" ]; then
    echo "Error: $CLASP_JSON not found. Please run this script from the project root."
    exit 1
fi

# Iterate over each scriptId provided as an argument
for SCRIPT_ID in "$@"; do
    echo "===================================================="
    echo "Target Script ID: $SCRIPT_ID"

    # 1. Update scriptId in .clasp.json
    echo "Updating .clasp.json..."
    jq --arg id "$SCRIPT_ID" '.scriptId = $id' "$CLASP_JSON" > "${CLASP_JSON}.tmp" && mv "${CLASP_JSON}.tmp" "$CLASP_JSON"

    # 2. clasp push
    echo "Running: clasp push"
    clasp push
    if [ $? -ne 0 ]; then
        echo "Error: clasp push failed for $SCRIPT_ID. Skipping to next."
        continue
    fi

    # 3. Create a new version
    echo "Creating new version..."
    NEW_VERSION=$(clasp version "Auto-deploy $(date +'%Y-%m-%d %H:%M:%S')" --json | jq -r '.versionNumber')
    echo "New Version: $NEW_VERSION"

    # 4. Get the latest versioned deployment ID
    echo "Fetching deployments..."
    # We filter for the first deployment that has a versionNumber (usually the most recent versioned one)
    DEPLOYMENT_ID=$(clasp deployments --json | jq -r 'map(select(.versionNumber != null)) | .[0].deploymentId')

    if [ "$DEPLOYMENT_ID" == "null" ] || [ -z "$DEPLOYMENT_ID" ]; then
        # Fallback: if no versionNumber found, try the first one
        DEPLOYMENT_ID=$(clasp deployments --json | jq -r '.[0].deploymentId')
    fi

    if [ "$DEPLOYMENT_ID" != "null" ] && [ -n "$DEPLOYMENT_ID" ]; then
        echo "Target Deployment ID: $DEPLOYMENT_ID"

        # 5. clasp redeploy {ID} {version}
        echo "Running: clasp redeploy $DEPLOYMENT_ID -V $NEW_VERSION"
        clasp redeploy "$DEPLOYMENT_ID" -V "$NEW_VERSION"

        # 6. Open Web App URL
        WEBAPP_URL="https://script.google.com/macros/s/$DEPLOYMENT_ID/exec"
        echo "Opening Web App: $WEBAPP_URL"
        xdg-open "$WEBAPP_URL"
    else
        echo "Warning: No deployments found for $SCRIPT_ID. Skip redeploy."
    fi

done

echo "===================================================="
echo "Done."
