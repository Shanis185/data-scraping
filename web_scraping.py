import pandas as pd
import random
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# Set up WebDriver
service = Service('C:\Users\shani\.wdm\drivers\chromedriver\win64\129.0.6668.100\chromedriver-win32/chromedriver.exe')
# Replace with your chromedriver path
driver = webdriver.Chrome(service=service, options=chrome_options)


def scroll_page():
    """Scroll the page to load more products"""
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)  # Wait for page to load
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


def scrape_products(url, min_products=200):
    products = []
    driver.get(url)

    while len(products) < min_products:
        try:
            # Scroll to load more products
            scroll_page()

            # Wait for products to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".product-card"))
            )

            # Find all product elements
            product_elements = driver.find_elements(By.CSS_SELECTOR, ".product-card")

            for index, product in enumerate(product_elements):
                if len(products) >= min_products:
                    break

                try:
                    # Extract product details
                    name = product.find_element(By.CSS_SELECTOR, ".sc-b74d844-17.fMLfKo").text.strip()

                    try:
                        price = product.find_element(By.CSS_SELECTOR, ".priceNow").text.strip().replace('AED',
                                                                                                        '').replace(',',
                                                                                                                    '').strip()
                    except NoSuchElementException:
                        price = 'N/A'

                    try:
                        old_price_element = product.find_element(By.CSS_SELECTOR, ".sc-14685a92-0.dROeqP")
                        old_price = old_price_element.text.strip().replace('AED', '').replace(',', '').strip()
                    except NoSuchElementException:
                        old_price = 'N/A'

                    try:
                        brand = product.find_element(By.CSS_SELECTOR, ".sc-b74d844-16.cNYrWp").text.strip()
                    except NoSuchElementException:
                        brand = 'N/A'

                    try:
                        seller = product.find_element(By.CSS_SELECTOR, ".sc-ae474822-2.eFjDyA").text.strip()
                    except NoSuchElementException:
                        seller = 'N/A'

                    try:
                        link = product.find_element(By.CSS_SELECTOR, "div.sc-19767e73-0.bwele a").get_attribute('href')
                    except NoSuchElementException:
                        link = 'N/A'

                    try:
                        rating_count = product.find_element(By.CSS_SELECTOR, ".sc-2709a77c-2.hUinXQ").text.strip()
                    except NoSuchElementException:
                        rating_count = 'N/A'

                    try:
                        average_rating = product.find_element(By.CSS_SELECTOR, ".sc-2709a77c-0.jHwyep").text.strip()
                    except NoSuchElementException:
                        average_rating = 'N/A'

                    try:
                        sku = product.get_attribute('id').split('-')[-1]
                    except:
                        sku = 'N/A'

                    try:
                        date_element = product.find_element(By.CSS_SELECTOR, ".sc-87126449-0.gIGZPG")
                        date_text = date_element.text.strip()
                        date = date_text.replace('Get it by', '').strip() if 'Get it by' in date_text else date_text
                    except NoSuchElementException:
                        date = 'N/A'

                    # Express delivery check
                    try:
                        express_element = product.find_element(By.CSS_SELECTOR, "div[class^='sc-d64b8217']")
                        is_express = 'Yes'
                    except NoSuchElementException:
                        is_express = 'No'

                    # Free delivery check
                    try:
                        free_delivery_element = product.find_element(By.CSS_SELECTOR, ".sc-1239b050-5")
                        has_free_delivery = 'Yes' if 'Free Delivery' in free_delivery_element.text else 'No'
                    except NoSuchElementException:
                        has_free_delivery = 'No'

                    # Append product data
                    products.append({
                        'Search Rank': len(products) + 1,
                        'Name': name,
                        'Price': float(price) if price != 'N/A' else 'N/A',
                        'Old Price': float(old_price) if old_price != 'N/A' else 'N/A',
                        'Brand': brand,
                        'Seller': seller,
                        'Link': link,
                        'Rating Count': rating_count,
                        'Average Rating': average_rating,
                        'SKU': sku,
                        'Date': date,
                        'Express': is_express,
                        'Free Delivery': has_free_delivery
                    })

                except Exception as e:
                    print(f"Error scraping product {len(products) + 1}: {str(e)}")
                    continue

            # Check if we need to load more products
            if len(products) < min_products:
                # Try to click "Load More" button if it exists
                try:
                    load_more = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, ".load-more-button"))
                    )
                    load_more.click()
                    time.sleep(2)  # Wait for new products to load
                except TimeoutException:
                    print(f"No more products to load. Scraped {len(products)} products.")
                    break

        except Exception as e:
            print(f"Error during scraping: {str(e)}")
            break

    return products


# URL of the product category
url = "https://www.noon.com/uae-en/sports-and-outdoors/exercise-and-fitness/yoga-16328/"

try:
    # Scrape products
    products = scrape_products(url)

    # Create DataFrame
    df = pd.DataFrame(products)

    # Save to CSV
    df.to_csv('yoga_products_selenium.csv', index=False)

    # Data Analysis
    most_expensive = df.loc[df['Price'].idxmax()] if 'Price' in df.columns else None
    cheapest = df.loc[df['Price'].idxmin()] if 'Price' in df.columns else None
    highest_ranked = df.loc[df['Search Rank'].idxmin()]
    lowest_ranked = df.loc[df['Search Rank'].idxmax()]
    products_per_brand = df['Brand'].value_counts()
    products_per_seller = df['Seller'].value_counts()
    free_delivery_count = df['Free Delivery'].value_counts()

    # Display results
    print(f"Total Products Scraped: {len(df)}")
    if most_expensive is not None:
        print("\nMost Expensive Product:")
        print(most_expensive)
    if cheapest is not None:
        print("\nCheapest Product:")
        print(cheapest)
    print("\nHighest Ranked Product:")
    print(highest_ranked)
    print("\nLowest Ranked Product:")
    print(lowest_ranked)
    print("\nNumber of Products from Each Brand:")
    print(products_per_brand)
    print("\nNumber of Products by Each Seller:")
    print(products_per_seller)
    print("\nNumber of Products with Free Delivery:")
    print(free_delivery_count)

    # Calculate correlation between rank and price
    if 'Price' in df.columns:
        numeric_df = df[df['Price'] != 'N/A'].copy()
        numeric_df['Price'] = pd.to_numeric(numeric_df['Price'])
        correlation = numeric_df['Search Rank'].corr(numeric_df['Price'])
        print(f"\nCorrelation between Search Rank and Price: {correlation:.2f}")

    # Plotting
    plt.figure(figsize=(10, 5))
    products_per_brand.plot(kind='bar', title='Number of Products per Brand', color='skyblue')
    plt.xlabel('Brand')
    plt.ylabel('Number of Products')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    plt.figure(figsize=(10, 5))
    products_per_seller.plot(kind='bar', title='Number of Products by Seller', color='lightgreen')
    plt.xlabel('Seller')
    plt.ylabel('Number of Products')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

    plt.figure(figsize=(8, 5))
    free_delivery_count.plot(kind='bar', title='Products with Free Delivery', color='salmon')
    plt.xlabel('Free Delivery')
    plt.ylabel('Number of Products')
    plt.xticks(rotation=0)
    plt.tight_layout()
    plt.show()

    # New plot for rank distribution
    plt.figure(figsize=(12, 6))
    df['Search Rank'].hist(bins=20, color='purple', alpha=0.7)
    plt.title('Distribution of Product Rankings')
    plt.xlabel('Search Rank')
    plt.ylabel('Number of Products')
    plt.tight_layout()
    plt.show()

    # Optional: Scatter plot of Rank vs Price if price data is available
    if 'Price' in df.columns:
        numeric_df = df[df['Price'] != 'N/A'].copy()
        numeric_df['Price'] = pd.to_numeric(numeric_df['Price'])
        plt.figure(figsize=(12, 6))
        plt.scatter(numeric_df['Search Rank'], numeric_df['Price'], alpha=0.5)
        plt.title('Search Rank vs Price')
        plt.xlabel('Search Rank')
        plt.ylabel('Price (AED)')
        plt.tight_layout()
        plt.show()

finally:
    # Close the browser
    driver.quit()