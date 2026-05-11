# Clean and Green Turf — Website

Single-page responsive website for **Clean and Green Turf Care and Renovation**.
Domain: **cleanandgreenturf.com**

## Files

| File | Purpose |
|---|---|
| `index.html` | All page content (hero, comparison, services, CTA, footer) |
| `styles.css` | All styles. Brand palette as CSS variables at top of file |
| `assets/` | Logo, future photos |

No JavaScript framework, no build step, no dependencies. Pure HTML + CSS.
Fonts are loaded from Google Fonts (Inter family).

## Deploying to Cloudflare Pages (recommended)

Since you (likely) registered the domain at Cloudflare, this is the simplest path.

1. Push this repo to GitHub (already done if you're reading this on GitHub).
2. Log in at **dash.cloudflare.com → Workers & Pages → Create application → Pages**.
3. **Connect to Git** → select this repository.
4. Build settings:
   - **Framework preset:** *None*
   - **Build command:** *(leave blank)*
   - **Build output directory:** `website`
5. Click **Save and Deploy**. First deploy takes ~30 seconds.
6. After it's live at `<random>.pages.dev`, go to **Custom domains → Set up a custom domain** and add `cleanandgreenturf.com` and `www.cleanandgreenturf.com`.
7. Cloudflare auto-configures the DNS since you own the domain there.

**Total cost: $0/month** (Cloudflare Pages is free for personal projects with unlimited bandwidth.)

## Deploying to Netlify (alternative)

1. Log in at app.netlify.com.
2. **Add new site → Import an existing project** → connect this repo.
3. **Publish directory:** `website` → Deploy.
4. **Domain settings → Add custom domain** → `cleanandgreenturf.com`.
5. Update the domain's DNS to point to Netlify (instructions Netlify provides).

## Editing copy

All content is in `index.html` between the section comments. Look for headers like:
```html
<!-- ────────────  HERO  ──────────── -->
```
Just edit the text between the tags. No need to touch CSS unless you're adjusting design.

## Brand palette (in CSS)

```css
--fairway: #2E7D32;  /* primary green */
--forest:  #1B4D1F;  /* dark green */
--fresh:   #A5D6A7;  /* light green */
--char:    #2C2C2C;  /* body text */
--lime:    #F4F4EF;  /* off-white background */
--yellow:  #FFC107;  /* accents only */
```

## What's still TODO

- [ ] Replace emoji icons in benefits/reasons with SVG icons (matches business card / door hanger style better than emoji)
- [ ] Add real before/after photos (currently uses gradient placeholders in CSS — see hero, services, about)
- [ ] Add Google Business Profile link in footer once profile is live
- [ ] Add Instagram / Facebook handles in footer once social is set up
- [ ] Consider adding a simple contact form (Cloudflare Pages supports Workers; or use Formspree free tier)
- [ ] Add favicon (use a square crop of the FINAL_logo_lockup)
- [ ] Add Open Graph share image (`/assets/og-image.jpg` at 1200x630)
- [ ] Schedule a Google PageSpeed audit before launch and fix any warnings

## Local preview

Open `index.html` in a browser directly — no server needed for basic preview. For mobile testing, run any static server:

```bash
cd website
python3 -m http.server 8000
# then visit http://localhost:8000 on your phone (same Wi-Fi network)
```
