# Excel Inventory Tracker

I built this for small online sellers who keep stock in a messy sheet and only find out something is out of stock when an order fails. It's a plain Excel file with formulas, alerts and a reorder dashboard. I generate it with Python so I can rebuild it quickly for each client.

[Watch the 38-second demo video](docs/demo.mp4)

![Dashboard](docs/step4_dashboard.png)

## What's inside

There are five sheets. How-To explains how to use it in simple steps. Products is the SKU list, and stock, status and days of stock left are worked out by formula. Stock In is where you log deliveries, and stock marked as incoming isn't counted until it arrives. Sales has a SKU dropdown and fills in price and revenue by itself. Dashboard shows the totals, a low-stock list and a stock vs reorder level chart.

Rows turn yellow when stock reaches the reorder level and red when it hits zero. There are no macros.

## How I tested it

I didn't want to show fake numbers, so I used real sales from the UCI Online Retail dataset, which comes from a UK gift shop. I took its 15 most-ordered products and their actual daily sales from 1 Nov to 9 Dec 2011. That's 704 rows and 57,598 units.

On 9 Dec 2011 the tracker showed 5 products out of stock, 3 low and 7 fine.

To be honest about one thing: the dataset only has sales. It has no stock or cost figures, so I had to assume those. I set opening stock on 1 Nov 2011 to 60 days of average daily demand (based on Dec 2010 to Oct 2011). The reorder level is 14 days of demand, and cost is 60% of the average selling price. So the stock levels are a simulation driven by real sales, not a real warehouse count.

## Try it

To see the demo, just open `inventory_tracker.xlsx`. To rebuild it:

```bash
pip install -r requirements.txt
python build_tracker.py                  # uses data/real_data.json (37 KB, in the repo)
# optional, to rebuild real_data.json from the original dataset:
python scripts/download_data.py && python scripts/prep_real_data.py
```

To use it for your own shop, delete the demo rows and type in your products, deliveries and sales.

## Screenshots

| | |
|---|---|
| ![Input](docs/step1_input.png) | ![Build](docs/step2_build.png) |
| ![Output](docs/step3_output.png) | ![Dashboard](docs/step4_dashboard.png) |

## Data and licence

The test data is Chen, D. (2015), *Online Retail*, UCI Machine Learning Repository, https://doi.org/10.24432/C5BW33, licensed CC BY 4.0. `data/real_data.json` is my small extract of it, with the top 15 products and their daily totals. The raw file isn't in the repo.

The code is MIT licensed. See [LICENSE](LICENSE).

Made by Shannu, a mechanical engineering student who likes automating boring work.
