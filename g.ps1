#!/usr/bin/env pwsh

# Check if a commit message was provided
if (-not $args[0]) {
    Write-Host "Error: No commit message provided."
    Write-Host "Usage: ./g.ps1 'Your commit message'"
    exit 1
}

# Store the commit message
$commitMessage = $args[0]

# Add all changes with logging
Write-Host "Adding all changes..."
git add .

# Commit with the provided message with logging
Write-Host "Committing changes with message: '$commitMessage'"
git commit -m $commitMessage

# Push to the current branch with logging
Write-Host "Pushing to the current branch..."
git push

# Log completion message
Write-Host "Changes have been successfully pushed." 