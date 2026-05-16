# Cosmic Pig Smoke Co.

Out-of-this-world smoke — small-batch, handcrafted BBQ rubs.

This repo holds the one-page site for the brand.

## Run it

It's a single self-contained file. Open `index.html` in a browser, or serve the folder:

```
python3 -m http.server 8000
# then visit http://localhost:8000
```

## Assets

- `assets/logo.png` — brand logo (hero + footer)
- `assets/labels/*.png` — one wraparound label per rub

The site auto-detects these files: if an image exists at the expected path
it replaces the styled placeholder automatically — no markup changes needed.

## The lineup

| Rub | Type |
| --- | --- |
| Solar Flare | Sweet Heat BBQ Rub (flagship) |
| Cosmic Wingman | Citrus Chile Chicken Rub |
| Black Hole Brisket | Texas-Style Brisket Rub |
| Moo-Nar Landing | Premium Beef Rub |
| Dark Matter | Activated Charcoal Beef Rub |
