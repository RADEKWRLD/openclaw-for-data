"""Base class for city-specific scrapers."""

from __future__ import annotations

import time
from abc import ABC, abstractmethod
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from ..models import PriceData


class BaseCityScraper(ABC):
    """Abstract base class for city-specific price scrapers.

    To add a new city, create a subclass and implement:
    - CITY_NAME: display name (e.g., "广州")
    - SOURCE_NAME: data source name
    - get_download_page_url(): return the page URL containing download links
    - find_excel_links(): parse a page to find Excel download links
    - parse_price_page(): parse HTML into PriceData (for HTML-based sites)

    Then register in cities/__init__.py.
    """

    CITY_NAME: str = ""
    SOURCE_NAME: str = ""
    BASE_URL: str = ""
    REQUEST_DELAY: float = 2.0
    REQUEST_TIMEOUT: int = 30
    USER_AGENT: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.USER_AGENT,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        })

    @abstractmethod
    def get_download_page_url(self, year: int, month: int) -> str:
        """Return URL of the page containing download links for a given month.

        Args:
            year: Target year.
            month: Target month (1-12).

        Returns:
            URL string.
        """
        ...

    @abstractmethod
    def find_excel_links(self, html: str, year: int, month: int) -> list[str]:
        """Parse download page HTML to find Excel file download URLs.

        Args:
            html: Page HTML content.
            year: Target year.
            month: Target month.

        Returns:
            List of download URLs for Excel files.
        """
        ...

    def fetch_page(self, url: str) -> str:
        """Fetch a page with retry logic."""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                resp = self.session.get(url, timeout=self.REQUEST_TIMEOUT)
                resp.raise_for_status()
                if resp.apparent_encoding:
                    resp.encoding = resp.apparent_encoding
                return resp.text
            except requests.RequestException as e:
                if attempt < max_retries - 1:
                    wait = (attempt + 1) * 2
                    print(f"  [retry] 请求失败 ({e}), {wait}s 后重试...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"请求失败 ({url}): {e}") from e
        return ""

    def download_file(self, url: str, save_path: Path) -> Path:
        """Download a file (Excel/PDF) from URL to local path."""
        save_path.parent.mkdir(parents=True, exist_ok=True)
        max_retries = 3
        for attempt in range(max_retries):
            try:
                resp = self.session.get(url, timeout=60, stream=True)
                resp.raise_for_status()
                with open(save_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f"  [download] 已下载: {save_path.name} ({save_path.stat().st_size // 1024} KB)")
                return save_path
            except requests.RequestException as e:
                if attempt < max_retries - 1:
                    wait = (attempt + 1) * 2
                    print(f"  [retry] 下载失败 ({e}), {wait}s 后重试...")
                    time.sleep(wait)
                else:
                    raise RuntimeError(f"下载失败 ({url}): {e}") from e
        return save_path

    def download_excels(
        self,
        year: int,
        months: list[int],
        save_dir: Path,
    ) -> dict[str, Path]:
        """Download Excel files for the given months.

        Args:
            year: Target year.
            months: Month numbers.
            save_dir: Directory to save downloaded files.

        Returns:
            Dict mapping month string to downloaded file path.
        """
        save_dir.mkdir(parents=True, exist_ok=True)
        downloaded: dict[str, Path] = {}

        for m in months:
            month_str = f"{year}-{m:02d}"
            print(f"\n[scrape] {month_str}: 查找下载链接...")

            try:
                page_url = self.get_download_page_url(year, m)
                html = self.fetch_page(page_url)
                links = self.find_excel_links(html, year, m)

                if not links:
                    print(f"  [warning] {month_str}: 未找到 Excel 下载链接")
                    continue

                # Download first matching link
                url = links[0]
                ext = ".xls" if ".xls" in url.lower() and ".xlsx" not in url.lower() else ".xlsx"
                filename = f"{year}-{m:02d}_信息价{ext}"
                save_path = save_dir / filename

                # Skip if already downloaded
                if save_path.exists():
                    print(f"  [cache] 已存在: {filename}")
                    downloaded[month_str] = save_path
                else:
                    downloaded[month_str] = self.download_file(url, save_path)

            except Exception as e:
                print(f"  [error] {month_str}: {e}")

            time.sleep(self.REQUEST_DELAY)

        return downloaded

    # Legacy methods for backward compatibility
    def get_monthly_urls(self, year: int, months: list[int]) -> dict[str, str]:
        urls = {}
        for m in months:
            month_str = f"{year}-{m:02d}"
            urls[month_str] = self.get_download_page_url(year, m)
        return urls

    def parse_price_page(self, html: str, month: str) -> PriceData:
        raise NotImplementedError("此城市插件使用 download_excels 模式，不支持 HTML 解析")
