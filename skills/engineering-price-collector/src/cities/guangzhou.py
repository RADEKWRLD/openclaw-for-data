"""Guangzhou (广州) city scraper for construction price information.

Data source: 广州市建设工程造价管理站
Website: https://gc.gzcc.gov.cn
"""

from __future__ import annotations

import re

from bs4 import BeautifulSoup

from .base import BaseCityScraper


class GuangzhouScraper(BaseCityScraper):
    CITY_NAME = "广州"
    SOURCE_NAME = "广州市建设工程造价管理站"
    BASE_URL = "https://gc.gzcc.gov.cn"

    def get_download_page_url(self, year: int, month: int) -> str:
        """Return URL for Guangzhou price bulletin download page.

        The actual URL pattern depends on the current site structure.
        This may need adjustment based on site changes.
        """
        return f"{self.BASE_URL}/xxj/list?year={year}&month={month}"

    def find_excel_links(self, html: str, year: int, month: int) -> list[str]:
        """Find Excel download links from Guangzhou price bulletin page.

        Looks for links containing .xls or .xlsx in href, matching the
        target year/month in link text or URL.
        """
        soup = BeautifulSoup(html, "lxml")
        links = []

        month_keywords = [
            f"{year}年{month}月",
            f"{year}-{month:02d}",
            f"{year}{month:02d}",
            "信息价",
        ]

        for a in soup.find_all("a", href=True):
            href = a["href"]
            text = a.get_text(strip=True)

            # Check if link points to an Excel file
            is_excel = any(ext in href.lower() for ext in [".xls", ".xlsx", ".et"])

            # Check if link text or URL matches our target period
            matches_period = any(kw in text or kw in href for kw in month_keywords)

            if is_excel and matches_period:
                # Resolve relative URLs
                if href.startswith("/"):
                    href = f"{self.BASE_URL}{href}"
                elif not href.startswith("http"):
                    href = f"{self.BASE_URL}/{href}"
                links.append(href)

        # Also look for download buttons or data-url attributes
        for elem in soup.find_all(attrs={"data-url": True}):
            url = elem["data-url"]
            if any(ext in url.lower() for ext in [".xls", ".xlsx"]):
                if url.startswith("/"):
                    url = f"{self.BASE_URL}{url}"
                links.append(url)

        return links
