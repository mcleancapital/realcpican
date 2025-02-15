import feedparser
import requests
from bs4 import BeautifulSoup
from datetime import datetime

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
            margin-bottom: 15px;
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
            max-width: 250px;
            height: auto;
            display: block;
            margin-bottom: 10px;
        }

        @media (max-width: 729px) {
            .mobile-menu {
                display: flex;
                justify-content: space-around;
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                background-color: #ffffff;
                border-bottom: 1px solid #ccc;
                z-index: 1000;
                padding: 10px 0;
            }
            .mobile-menu button {
                flex: 1;
                padding: 10px;
                font-size: 16px;
                border: none;
                background-color: #fff;
                cursor: pointer;
                transition: background-color 0.3s ease;
            }
            .mobile-menu button:hover {
                background-color: #e0e0e0;
            }
            body {
                margin-top: 50px;
            }
        }

        footer {
            padding: var(--std_padding);
            box-shadow: rgb(0 0 0 / 8%) 0 -1px 0;
            text-align: center;
            color: #717171;
            font-size: .75em;
        }

        @media(min-width:730px) {
            .mobile-menu {
                display: none;
            }
            footer {
                display: none;
            }
        }
    </style>
</head>
<body>
<div class="mobile-menu">
    <img src="/static/images/logo.jpg" alt="Real CPI Logo" height="30" />
    <a href="https://www.realcpi.org"><button id="data-button">Data</button></a>
    <a href="https://www.realcpi.org/news.html"><button id="news-button">News</button></a>
</div>
<br>
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
for entry in feed.entries[:30]:  # Fetch the latest 30 articles
    # Get the article URL
    article_url = entry.link

    # Fetch the image URL
    image_url = fetch_image_url(article_url)

    # Ensure the image URL is properly formatted to prevent '&' encoding issues
    image_url = image_url.replace("&", "&amp;") if image_url else None

    # Debug: Print the modified image URL
    print(f"Modified Image URL: {image_url}")

    # Use the `get` method to safely access the `summary` attribute with a default value
    summary = entry.get('summary', 'No summary available.')

    # Add the article to the HTML, including the image if available
    html_content += f"""
        <li>
            {"<img src='" + image_url + "' alt='Article Image'>" if image_url else ""}
            <h3>
                <a href="{article_url}" target="_blank" rel="noopener noreferrer"><b>{entry.title}</b></a>
                <p><i>Source:</i> CNBC</p>
            </h3>
            <p>{summary}</p>
        </li>
    """

# Close the HTML structure
html_content += """
        </ul>
    </div>

<footer>
This information is provided for informational purposes only. This is not for advice.
<br><br>
Please let us know if you have any suggestions, including data series or metrics we could add to Real CPI.
<br><br>
Copyright © 2024
&nbsp;
&nbsp;
&nbsp;
&nbsp;
<a href="mailto:contact@realcpi.org">contact@realcpi.org</a>
</footer>

</body>
</html>
"""

# Save the HTML file
output_file = "news_mobile.html"
with open(output_file, "w", encoding="utf-8") as file:
    file.write(html_content)

print(f"{output_file} has been updated with the latest news from CNBC.")

# Save the second file with a timestamp in the "news-archive" folder
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_file_archive = f"news_{timestamp}.html"
with open(output_file_archive, "w", encoding="utf-8") as file:
    file.write(html_content)

print(f"Archive file {output_file_archive} has been created.")
