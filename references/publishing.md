# Publish the animated deck

Use Vercel or Here.now according to the user's request and the available authorized route. Discover the current provider tools and read their instructions; do not assume CLIs, credentials, deployment APIs, pricing or free limits exist. A private source repository does not imply a private deployed website.

Before deploying, inspect the exact directory and intended audience. Authorize public exposure of the specific content as required by current policy. User-provided confidential documents and sensitive personal information require extra care. Publish only the reviewed HTML and needed assets; do not include source reports, contacts lists, credentials, raw prompts or internal notes. Do not enable public downloads of the PPTX or bonus PDFs unless that sharing is also authorized.

Use an existing authenticated provider connection when supported. If login, new account terms, payment or security access is required, request the specific missing step rather than inventing credentials or silently switching platforms. Do not spend money or upgrade a plan without approval.

After deployment, wait until the build is complete. Copy the exact URL returned by the provider, open it afresh and check the title, expected slide count, assets and navigation. Check both desktop and mobile. Confirm access for the intended recipient. Private preview URLs requiring the assistant's login are not successful delivery to the user.

If provider access is unavailable or authorization is pending, deliver completed local outputs and the portable HTML bundle, explain that only publishing remains blocked, and ask for the smallest next step. Never imply a local HTML file is hosted or claim a URL was verified when only the deployment API succeeded.

## Provider-specific checks

For Vercel, prefer the existing authorized connection and a verified account/project. Use current deployment/upload/status tools, wait for READY, then verify audience access anonymously. Do not disable Deployment Protection or generate bypass credentials simply to expose a URL. Official guidance: https://vercel.com/docs/deployments and https://vercel.com/docs/deployment-protection .

For Here.now, consult https://here.now/docs before use. Its direct publish flow currently posts a path/size/MIME manifest to /api/v1/publish, uploads to the returned upload URLs, then finalizes using the returned finalizeUrl and versionId, with the required X-HereNow-Client header. Follow returned endpoints exactly. Anonymous publishing currently expires after 24 hours: disclose expiry and return any exact claimUrl privately, never embed it into the public site. For authenticated persistence, inspect publishStatus; do not generate or store API keys automatically. Treat these facts as time-sensitive and recheck current provider instructions.
