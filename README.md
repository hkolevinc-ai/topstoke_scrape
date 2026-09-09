# TopStokee → Temu scraper

The scraper scans all product sitemaps on `topstokee.com`, reads the CloudCart product data, and populates the supplied Temu template.

## What it does

- scans the whole site in `full` mode;
- exports each sellable variation on a separate row;
- splits mixed adult and children's sizes into separate Temu parent products;
- uses the live discounted price and keeps the old price as the list price when available;
- collects up to 10 SKU images and available detail/size-guide images;
- removes the promotional delivery/payment/exchange text from descriptions;
- splits output automatically into files of at most 1,900 data rows;
- produces a raw CSV, a skipped-products CSV, a log, and a summary.

## GitHub Actions

1. Upload every file and folder from this package to a new GitHub repository. Keep `.github/workflows/scraper.yml` in the same path.
2. Open **Actions → TopStokee Temu scraper → Run workflow**.
3. First choose `test`, leave `max_products` at `0`, and run it. Test mode scans 15 products.
4. Download the `topstokee-temu-results` artifact and test one generated XLSX in Temu.
5. After the test succeeds, run again with `mode = full` and `max_products = 0`.

A full run currently discovers about 2,300 product pages and can take roughly 45–90 minutes, depending on the site's response time. GitHub Actions keeps the generated files as a downloadable artifact.

## Assumptions to review in `config.json`

- quantity is `10` for every sellable variation because the site does not publish exact stock quantities;
- shipping template is `OFIS`, taken from the supplied Temu file;
- manufacturer is `TOP STOKE EOOD`, taken from the supplied Temu file;
- country of origin is set to `Bulgaria`;
- color is `Multicolor` when CloudCart does not publish a color variation;
- package weight/dimensions and fallback fabric composition use product-type defaults in `scraper.py`.

Products for which the supplied Temu template has no suitable category, such as standalone sweatpants or backpacks, are not forced into an incorrect category. They are listed in `topstokee_skipped_products.csv`.

## Local run

```bash
python -m pip install -r requirements.txt
python scraper.py --mode test
python scraper.py --mode full --max-products 0
```
