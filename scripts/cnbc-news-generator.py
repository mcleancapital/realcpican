import feedparser
from bs4 import BeautifulSoup
import requests

# CNBC Business News RSS Feed URL
RSS_FEED_URL = "https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=10001147"

# Parse the RSS feed
feed = feedparser.parse(RSS_FEED_URL)

# Start building the HTML content
html_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Top Business News</title>
    <style>
        body {
            font-family: garamond, sans-serif;
            line-height: 1.4;
            margin: 15px;
            padding: 0;
        }
        h2 {
            font-weight: bold;
            text-align: left;
        }
        a {
            color: green;
            text-decoration: none;
        }
        a:hover {
            text-decoration: underline;
        }
        ul {
            padding: 0;
            list-style: none;
        }
        li {
            margin-bottom: 30px;
        }
        p {
            margin: 0;
            font-size: 16px;
        }
        h3 {
            font-size: 16px;
            font-weight: normal;
        }
        img {
            max-width: 250px; /* Set a maximum width for images */
            height: auto; /* Maintain aspect ratio */
            display: block;
            margin-bottom: 10px; /* Add spacing below the image */
        }
    </style>
</head>
<body>
    <div id="news-section">
        <h2>Top Business News</h2>
        <ul>
"""

# Function to fetch the image URL from the article page
def fetch_image_url(article_url):
    try:
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        }
        response = requests.get(article_url, headers=headers)

        # Check response status
        print(f"Fetching {article_url} - Status Code: {response.status_code}")
        if response.status_code != 200:
            return None

        soup = BeautifulSoup(response.content, "html.parser")
        # Look for the Open Graph meta tag with property="og:image"
        meta_image = soup.find("meta", property="og:image")
        if meta_image:
            return meta_image["content"]
    except Exception as e:
        print(f"Error fetching image for {article_url}: {e}")
    return None

# Loop through the articles and add them to the HTML
for entry in feed.entries[:15]:  # Fetch the latest 15 articles
    # Get the article URL
    article_url = entry.link

    # Fetch the image URL
    image_url = fetch_image_url(article_url)

    # Ensure the image URL is properly formatted to prevent '&' encoding issues
    image_url = image_url.replace("&", "&amp;") if image_url else None

    # Debug: Print the modified image URL
    print(f"Modified Image URL: {image_url}")

    # Add the article to the HTML, including the image if available
    html_content += f"""
        <li>
            {"<img src='" + image_url + "' alt='Article Image'>" if image_url else ""}
            <h3>
                <a href="{article_url}" target="_blank" rel="noopener noreferrer"><b>{entry.title}</b></a>
                <p><i>Source:</i> CNBC</p>
            </h3>
            <p>{entry.summary}</p>
        </li>
    """

# Close the HTML structure
html_content += """
        </ul>
    </div>
</body>
</html>
"""

# Save the HTML file
output_file = "news.html"
with open(output_file, "w", encoding="utf-8") as file:
    file.write(html_content)

print(f"{output_file} has been updated with the latest news from CNBC.")
