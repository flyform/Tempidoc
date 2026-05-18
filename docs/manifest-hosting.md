# Manifest & Hosting

## Manifest requirements

- The manifest must specify **WordApi 1.9** minimum.

## Hosting model

- The add-in must be hosted as static files from a stable web endpoint.
- Typical setup could be:
  - Nginx in a Docker container on Ubuntu with public access via Cloudflare Tunnel
  - Hosting via Cloudflare Pages
  - Hosting via GitHub Pages
- Avoid dynamic server requirements unless explicitly requested.

## Deployment notes

- **Manifest updates** require re-upload / redeploy through the M365 admin workflow used for Centralized Deployment.
