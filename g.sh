#!/bin/bash

# Check if a commit message was provided
if [ -z "$1" ]; then
  echo "Error: No commit message provided."
  echo "Usage: ./git_push.sh \"Your commit message\""
  exit 1
fi

# Add all changes with logging
echo "Adding all changes..."
git add .

# Commit with the provided message with logging
echo "Committing changes with message: '$1'"
git commit -m "$1"

# Push to the current branch with logging
echo "Pushing to the current branch..."
git push

# Log completion message
echo "Changes have been successfully pushed."
