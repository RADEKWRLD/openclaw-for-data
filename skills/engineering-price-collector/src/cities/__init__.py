from .base import BaseCityScraper
from .guangzhou import GuangzhouScraper

CITY_SCRAPERS: dict[str, type[BaseCityScraper]] = {
    "广州": GuangzhouScraper,
}


def get_scraper(city: str) -> BaseCityScraper:
    scraper_cls = CITY_SCRAPERS.get(city)
    if scraper_cls is None:
        supported = ", ".join(CITY_SCRAPERS.keys())
        raise ValueError(f"不支持的城市: {city}。当前支持: {supported}")
    return scraper_cls()
