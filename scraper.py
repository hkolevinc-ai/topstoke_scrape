#!/usr/bin/env python3
"""Scrape TopStokee and populate the supplied Temu bulk-upload template."""

from __future__ import annotations

import argparse
import csv
import html
import json
import logging
import random
import re
import shutil
import sys
import threading
import time
import unicodedata
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

from curl_cffi import requests as browser_requests
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


SITE_URL = "https://topstokee.com"
ROOT_SITEMAP = f"{SITE_URL}/sitemap.xml"
TEMPLATE_SHEET = "Template"
FIRST_DATA_ROW = 5
REQUEST_MIN_INTERVAL_SECONDS = 0.22

_HTTP_STATE = threading.local()
_RATE_LOCK = threading.Lock()
_NEXT_REQUEST_AT = 0.0
_AJAX_NOTICE_LOCK = threading.Lock()
_AJAX_FALLBACK_NOTICE_SHOWN = False

PROMO_PHRASES = (
    "безплатна доставка над 120лв",
    "безплатна доставка над 120 лв",
    "възможност за замяна",
    "доставка до 2 работни дни",
    "плати с карта",
    "дреха със собствен дизайн",
)

PROMO_IMAGE_MARKERS = (
    "teniska-65632283aa8a4",
    "free-delivery",
    "bezplatna-dostavka",
    "benefits",
)

MATERIAL_NAMES = {
    "acrylic": "Acrylic",
    "акрил": "Acrylic",
    "cashmere": "Cashmere",
    "кашемир": "Cashmere",
    "cotton": "Cotton",
    "памук": "Cotton",
    "памуч": "Cotton",
    "linen": "Linen",
    "лен": "Linen",
    "modal": "Modal",
    "модал": "Modal",
    "nylon": "Nylon",
    "найлон": "Nylon",
    "polyamide": "Polyamide",
    "полиамид": "Polyamide",
    "polyester": "Polyester",
    "полиестер": "Polyester",
    "silk": "Silk",
    "коприна": "Silk",
    "wool": "Wool",
    "вълна": "Wool",
    "elastane": "Elastane",
    "еластан": "Elastane",
    "lycra": "Elastane",
    "ликра": "Elastane",
    "viscose": "Viscose",
    "вискоза": "Viscose",
    "polyurethane": "Polyurethane",
    "полиуретан": "Polyurethane",
}

SUPPORTED_CATEGORY_IDS = {
    "tshirt_adult": "30469",
    "tshirt_kids": "30843",
    "hoodie_adult": "30467",
    "hoodie_kids": "30847",
    "sweatshirt_adult": "30466",
    "sweatshirt_kids": "30847",
    "tank_adult": "30475",
    "tank_kids": "30841",
    "vest_adult": "30407",
    "vest_kids": "30763",
    "windbreaker_adult": "30406",
    "windbreaker_kids": "30763",
    "shorts_adult": "30425",
    "shorts_kids": "30728",
    "short_set_adult": "27209",
    "long_set_adult": "30443",
    "set_kids": "30811",
    "hat": "30611",
    "polo_adult": "30472",
    "polo_kids": "34988",
}

DEFAULT_PACKAGE = {
    "tshirt": (300.0, 30.0, 25.0, 3.0),
    "hoodie": (750.0, 36.0, 30.0, 9.0),
    "sweatshirt": (650.0, 35.0, 30.0, 8.0),
    "tank": (250.0, 30.0, 24.0, 3.0),
    "vest": (650.0, 36.0, 30.0, 8.0),
    "windbreaker": (600.0, 36.0, 30.0, 7.0),
    "shorts": (400.0, 32.0, 27.0, 5.0),
    "short_set": (650.0, 36.0, 30.0, 8.0),
    "long_set": (950.0, 40.0, 32.0, 11.0),
    "set": (750.0, 36.0, 30.0, 9.0),
    "hat": (200.0, 24.0, 20.0, 12.0),
    "polo": (350.0, 32.0, 27.0, 4.0),
}

TOP_MEASUREMENTS_TSHIRT_ADULT = {
    "XXS": (42.0, 65.0),
    "XS": (45.0, 67.0),
    "S": (48.0, 69.0),
    "M": (50.0, 71.0),
    "L": (54.0, 73.0),
    "XL": (58.0, 75.0),
    "XXL": (62.0, 77.0),
    "XXXL": (66.0, 79.0),
}

TOP_MEASUREMENTS_TSHIRT_WOMEN = {
    "S": (41.0, 63.0),
    "M": (44.0, 64.0),
    "L": (47.0, 68.5),
    "XL": (50.0, 66.0),
    "XXL": (53.0, 67.0),
}

TOP_MEASUREMENTS_TSHIRT_KIDS = {
    "104": (38.0, 45.0),
    "116": (40.5, 50.0),
    "128": (43.0, 55.0),
    "140": (46.0, 60.0),
    "152": (48.5, 65.0),
}

TOP_MEASUREMENTS_HOODIE_ADULT = {
    "S": (54.0, 64.0),
    "M": (57.0, 67.0),
    "L": (60.0, 70.0),
    "XL": (63.0, 73.0),
    "XXL": (66.0, 76.0),
    "XXXL": (69.0, 79.0),
}

TOP_MEASUREMENTS_HOODIE_KIDS = {
    "104": (36.0, 44.0),
    "116": (40.0, 48.0),
    "128": (44.0, 52.0),
    "140": (46.0, 56.0),
    "152": (49.0, 60.0),
}

TOP_MEASUREMENTS_TANK_ADULT = {
    "S": (46.0, 68.0),
    "M": (48.5, 70.0),
    "L": (53.5, 72.0),
    "XL": (56.0, 74.0),
    "XXL": (58.5, 76.0),
    "XXXL": (60.5, 78.0),
}

TOP_MEASUREMENTS_WINDBREAKER_ADULT = {
    "S": (54.0, 65.0),
    "M": (57.0, 68.0),
    "L": (60.0, 71.0),
    "XL": (63.0, 74.0),
    "XXL": (66.0, 77.0),
}

TOP_MEASUREMENTS_VEST_ADULT = {
    "S": (52.0, 66.0),
    "M": (55.0, 68.0),
    "L": (58.0, 70.0),
    "XL": (62.0, 72.0),
    "XXL": (66.0, 73.0),
    "XXXL": (69.0, 75.0),
}

BOTTOM_MEASUREMENTS_SHORTS_ADULT = {
    "S": (38.5, 44.0),
    "M": (40.5, 46.0),
    "L": (42.5, 48.0),
    "XL": (44.5, 50.0),
    "XXL": (46.5, 52.0),
}

BOTTOM_MEASUREMENTS_SET_ADULT = {
    "S": (34.0, 43.0),
    "M": (35.0, 45.0),
    "L": (36.0, 48.0),
    "XL": (37.0, 51.0),
    "XXL": (38.0, 54.0),
}

BOTTOM_MEASUREMENTS_SET_KIDS = {
    "104": (21.0, 28.0),
    "116": (23.0, 30.0),
    "128": (26.0, 32.0),
    "140": (28.0, 34.0),
    "152": (30.0, 36.0),
}

SIZE_ALIASES = {
    "2XS": "XXS",
    "2XL": "XXL",
    "3XL": "XXXL",
}


@dataclass
class Config:
    site_url: str = SITE_URL
    root_sitemap: str = ROOT_SITEMAP
    default_quantity: int = 10
    max_rows_per_file: int = 1900
    workers: int = 3
    request_timeout: int = 45
    request_delay_seconds: float = 0.35
    shipping_template: str = "OFIS"
    manufacturer: str = "TOP STOKE EOOD"
    eu_responsible_person: str = ""
    country_of_origin: str = "Bulgaria"
    default_color: str = "Multicolor"
    default_composition: dict[str, dict[str, float]] = field(
        default_factory=lambda: {
            "tshirt": {"Cotton": 100.0},
            "hoodie": {"Cotton": 80.0, "Polyester": 20.0},
            "sweatshirt": {"Cotton": 80.0, "Polyester": 20.0},
            "tank": {"Cotton": 100.0},
            "vest": {"Cotton": 80.0, "Polyester": 20.0},
            "windbreaker": {"Polyester": 100.0},
            "shorts": {"Cotton": 80.0, "Polyester": 20.0},
            "short_set": {"Cotton": 80.0, "Polyester": 20.0},
            "long_set": {"Cotton": 80.0, "Polyester": 20.0},
            "set": {"Cotton": 80.0, "Polyester": 20.0},
            "hat": {"Polyester": 100.0},
            "polo": {"Cotton": 100.0},
        }
    )

    @classmethod
    def from_json(cls, path: Path | None) -> "Config":
        cfg = cls()
        if not path or not path.exists():
            return cfg
        data = json.loads(path.read_text(encoding="utf-8"))
        for key, value in data.items():
            if hasattr(cfg, key):
                setattr(cfg, key, value)
        return cfg


@dataclass
class Product:
    url: str
    product_id: str
    name: str
    category_path: str
    description: str
    bullets: list[str]
    main_images: list[str]
    detail_images: list[str]
    price: float
    list_price: float | None
    variants: list[dict[str, Any]]


@dataclass
class OutputRow:
    source_url: str
    source_category: str
    category_id: str
    category_key: str
    product_id: str
    parent_sku: str
    sku: str
    name: str
    description: str
    bullets: list[str]
    main_images: list[str]
    detail_images: list[str]
    size_raw: str
    size_family: str
    sub_size_family: str
    temu_size: str
    color: str
    price: float
    list_price: float | None
    quantity: int
    composition: dict[str, float]
    product_kind: str
    is_kids: bool
    age_group: str


def browser_session() -> Any:
    session = getattr(_HTTP_STATE, "session", None)
    if session is None:
        session = browser_requests.Session(impersonate="chrome")
        _HTTP_STATE.session = session
        _HTTP_STATE.warmed_hosts = set()
    return session


def wait_for_request_slot() -> None:
    global _NEXT_REQUEST_AT
    with _RATE_LOCK:
        now = time.monotonic()
        wait_seconds = max(0.0, _NEXT_REQUEST_AT - now)
        _NEXT_REQUEST_AT = max(now, _NEXT_REQUEST_AT) + REQUEST_MIN_INTERVAL_SECONDS
    if wait_seconds:
        time.sleep(wait_seconds)


def unwrap_ajax_body(source: str) -> str:
    stripped = source.lstrip()
    if not stripped.startswith("{"):
        return source
    try:
        payload = json.loads(source)
    except (TypeError, ValueError):
        return source
    body = payload.get("body") if isinstance(payload, dict) else None
    return body if isinstance(body, str) else source


def get_text(
    url: str,
    timeout: int,
    *,
    referer: str = "",
    ajax: bool = False,
    retry_forbidden: bool = False,
) -> str:
    headers = {
        "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7",
        "Cache-Control": "no-cache",
    }
    if referer:
        headers["Referer"] = referer
    if ajax:
        headers.update(
            {
                "Accept": "application/json, text/javascript, */*; q=0.01",
                "X-Requested-With": "XMLHttpRequest",
            }
        )
    last_error: Exception | None = None
    for attempt in range(5):
        wait_for_request_slot()
        try:
            response = browser_session().get(url, headers=headers, timeout=timeout, allow_redirects=True)
            if 200 <= response.status_code < 300:
                return unwrap_ajax_body(response.text) if ajax else response.text
            last_error = RuntimeError(f"HTTP {response.status_code}: {response.reason}")
            retryable = {429, 500, 502, 503, 504}
            if retry_forbidden:
                retryable.add(403)
            if response.status_code not in retryable:
                raise last_error
        except Exception as exc:
            last_error = exc
            status = getattr(getattr(exc, "response", None), "status_code", None)
            if status and status not in {403, 429, 500, 502, 503, 504}:
                raise
        if attempt < 4:
            time.sleep((0.9 * (2 ** attempt)) + random.uniform(0.15, 0.65))
    assert last_error is not None
    raise last_error


def warm_product_session(config: Config) -> None:
    session = browser_session()
    host = urlsplit(config.site_url).netloc.lower()
    warmed_hosts = getattr(_HTTP_STATE, "warmed_hosts", set())
    if host in warmed_hosts:
        return
    try:
        wait_for_request_slot()
        response = session.get(
            config.site_url.rstrip("/") + "/",
            headers={"Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7"},
            timeout=config.request_timeout,
            allow_redirects=True,
        )
        if response.status_code >= 400:
            logging.debug("Session warm-up returned HTTP %s", response.status_code)
    except Exception as exc:
        logging.debug("Session warm-up failed: %s", exc)
    warmed_hosts.add(host)
    _HTTP_STATE.warmed_hosts = warmed_hosts


def log_ajax_fallback_once() -> None:
    global _AJAX_FALLBACK_NOTICE_SHOWN
    with _AJAX_NOTICE_LOCK:
        if _AJAX_FALLBACK_NOTICE_SHOWN:
            return
        logging.warning("Regular product pages are blocked; using CloudCart's browser AJAX response")
        _AJAX_FALLBACK_NOTICE_SHOWN = True


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def product_urls_from_sitemap(source: str) -> list[str]:
    root = ET.fromstring(source)
    product_urls: list[str] = []
    for child in root.iter():
        if local_name(child.tag) == "loc" and child.text and "/product/" in child.text:
            product_urls.append(child.text.strip())
    return list(dict.fromkeys(product_urls))


def discover_from_root_sitemap(config: Config, max_urls: int = 0) -> list[str]:
    root = ET.fromstring(get_text(config.root_sitemap, config.request_timeout))
    sitemap_urls = [
        child.text.strip()
        for child in root.iter()
        if local_name(child.tag) == "loc" and child.text and "/sitemap/product/" in child.text
    ]
    product_urls: list[str] = []
    for sitemap_url in sitemap_urls:
        logging.info("Reading product sitemap: %s", sitemap_url)
        product_urls.extend(product_urls_from_sitemap(get_text(sitemap_url, config.request_timeout)))
        if max_urls and len(product_urls) >= max_urls:
            break
    return list(dict.fromkeys(product_urls))


def discover_from_direct_sitemaps(config: Config, max_urls: int = 0) -> list[str]:
    """Try CloudCart's predictable product sitemap URLs without reading the index."""
    product_urls: list[str] = []
    for index in range(1, 51):
        sitemap_url = urljoin(config.site_url, f"/sitemap/product/{index}.xml")
        try:
            logging.info("Trying direct product sitemap: %s", sitemap_url)
            batch = product_urls_from_sitemap(get_text(sitemap_url, config.request_timeout))
        except Exception as exc:
            if index == 1:
                logging.warning("Direct product sitemaps are unavailable: %s", exc)
            break
        if not batch:
            break
        product_urls.extend(batch)
        if max_urls and len(product_urls) >= max_urls:
            break
        # CloudCart currently writes at most 1,000 products per sitemap.
        if len(batch) < 1000:
            break
    return list(dict.fromkeys(product_urls))


def extract_catalogue_links(source: str, base_url: str) -> tuple[list[str], int]:
    products: list[str] = []
    for raw_href in re.findall(r"\bhref\s*=\s*['\"]([^'\"]+)", source, flags=re.IGNORECASE):
        absolute = urljoin(base_url, html.unescape(raw_href))
        parts = urlsplit(absolute)
        if parts.netloc.lower() == urlsplit(base_url).netloc.lower() and parts.path.startswith("/product/"):
            products.append(urlunsplit((parts.scheme, parts.netloc, parts.path.rstrip("/"), "", "")))
    page_counts = [int(value) for value in re.findall(r"data-pages=['\"](\d+)", source, flags=re.IGNORECASE)]
    return list(dict.fromkeys(products)), max(page_counts, default=1)


def discover_from_catalogue(config: Config, max_urls: int = 0) -> list[str]:
    """Fallback for hosts that block XML sitemap requests from GitHub runners."""
    catalogue_url = urljoin(config.site_url, "/category/produkti")
    logging.info("Using catalogue fallback: %s", catalogue_url)
    first_source = get_text(catalogue_url, config.request_timeout)
    product_urls, page_count = extract_catalogue_links(first_source, catalogue_url)
    if max_urls and len(product_urls) >= max_urls:
        return product_urls[:max_urls]

    pages = list(range(2, page_count + 1))
    if max_urls and product_urls:
        products_per_page = len(product_urls)
        required_pages = max(1, (max_urls + products_per_page - 1) // products_per_page)
        pages = pages[: max(0, required_pages - 1)]

    def fetch_page(page_number: int) -> tuple[int, list[str]]:
        page_url = f"{catalogue_url}?page={page_number}"
        source = get_text(page_url, config.request_timeout)
        links, _ = extract_catalogue_links(source, page_url)
        return page_number, links

    page_results: dict[int, list[str]] = {}
    with ThreadPoolExecutor(max_workers=config.workers) as executor:
        future_to_page = {executor.submit(fetch_page, page): page for page in pages}
        for future in as_completed(future_to_page):
            page = future_to_page[future]
            try:
                page_number, links = future.result()
                page_results[page_number] = links
                logging.info("Catalogue discovery page %d/%d", page_number, page_count)
            except Exception as exc:
                logging.warning("Catalogue page %d failed: %s", page, exc)
    for page in sorted(page_results):
        product_urls.extend(page_results[page])
    return list(dict.fromkeys(product_urls))[: max_urls or None]


def discover_product_urls(config: Config, max_urls: int = 0) -> list[str]:
    try:
        urls = discover_from_root_sitemap(config, max_urls)
        if urls:
            return urls[: max_urls or None]
    except Exception as exc:
        logging.warning("Root sitemap is unavailable; switching discovery method: %s", exc)

    urls = discover_from_direct_sitemaps(config, max_urls)
    if urls:
        return urls[: max_urls or None]

    urls = discover_from_catalogue(config, max_urls)
    if not urls:
        raise RuntimeError("No product URLs could be discovered from sitemap or catalogue")
    return urls


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def strip_accents(value: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", value) if not unicodedata.combining(ch))


def canonical_image_url(url: str, *, large: bool = True) -> str:
    url = html.unescape(url or "").strip()
    if not url:
        return ""
    parts = urlsplit(url)
    if not (large and "/cdn/img/products/" in parts.path):
        return urlunsplit((parts.scheme, parts.netloc, parts.path, parts.query, ""))
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query["width"] = "1920"
    query["height"] = "1920"
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query), ""))


def unique_strings(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        value = value.strip()
        if value and value not in seen:
            seen.add(value)
            result.append(value)
    return result


def clean_description_text(text: str) -> str:
    lines = []
    for raw_line in (text or "").replace("\xa0", " ").splitlines():
        line = normalize_space(raw_line)
        if not line:
            continue
        folded = normalize_space(strip_accents(line).lower())
        folded = re.sub(r"[^a-zа-я0-9]+", " ", folded).strip()
        if any(phrase in folded for phrase in PROMO_PHRASES):
            continue
        lines.append(line)
    cleaned = "\n".join(lines)
    for phrase in PROMO_PHRASES:
        cleaned = re.sub(re.escape(phrase), "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\bвизи\s+я\b", "визия", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip(" \n,;:-")
    return cleaned[:2000]


def extract_page_data(source: str) -> dict[str, Any]:
    match = re.search(r"window\.cc_page_data\s*=\s*(\{.*?\})\s*;", source, flags=re.DOTALL)
    if not match:
        raise ValueError("window.cc_page_data was not found")
    return json.loads(match.group(1))


class ProductHTMLExtractor(HTMLParser):
    VOID_TAGS = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"}

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.stack: list[tuple[str, str, set[str]]] = []
        self.description_container_depth: int | None = None
        self.textbox_depth: int | None = None
        self.skip_depth: int | None = None
        self.text_parts: list[str] = []
        self.detail_images: list[str] = []
        self.main_images: list[str] = []
        self.meta_description = ""
        self.current_bullet: list[str] | None = None
        self.bullets: list[str] = []

    def handle_starttag(self, tag: str, attrs_list: list[tuple[str, str | None]]) -> None:
        attrs = {key: value or "" for key, value in attrs_list}
        classes = set(attrs.get("class", "").split())
        self.stack.append((tag, attrs.get("id", ""), classes))
        depth = len(self.stack)

        if tag == "meta" and attrs.get("name", "").lower() == "description":
            self.meta_description = attrs.get("content", "")
        if tag == "div" and attrs.get("id") == "product-details-description":
            self.description_container_depth = depth
        if self.description_container_depth and tag == "div" and "_textbox" in classes:
            self.textbox_depth = depth
        if self.textbox_depth and tag in {"script", "style", "noscript"}:
            self.skip_depth = depth
        if self.textbox_depth and tag in {"p", "li"}:
            self.current_bullet = []
        if tag == "img":
            raw = attrs.get("data-large") or attrs.get("data-src") or attrs.get("src") or ""
            if self.textbox_depth:
                self.detail_images.append(raw)
            elif "primary" in classes or attrs.get("data-gallery-id"):
                self.main_images.append(raw)
        if tag in self.VOID_TAGS and self.stack:
            self.stack.pop()

    def handle_startendtag(self, tag: str, attrs_list: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs_list)
        if tag not in self.VOID_TAGS:
            self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        depth = len(self.stack)
        if self.current_bullet is not None and tag in {"p", "li"}:
            text = normalize_space(" ".join(self.current_bullet))
            if text:
                self.bullets.append(text)
            self.current_bullet = None
        if self.skip_depth == depth:
            self.skip_depth = None
        if self.textbox_depth == depth and tag == "div":
            self.textbox_depth = None
        if self.description_container_depth == depth and tag == "div":
            self.description_container_depth = None
        if self.stack:
            self.stack.pop()

    def handle_data(self, data: str) -> None:
        if self.textbox_depth and not self.skip_depth:
            text = normalize_space(data)
            if text:
                self.text_parts.append(text)
                if self.current_bullet is not None:
                    self.current_bullet.append(text)


def parse_product(url: str, config: Config) -> Product:
    warm_product_session(config)
    referer = urljoin(config.site_url, "/category/produkti")
    try:
        source = get_text(url, config.request_timeout, referer=referer)
    except Exception as regular_error:
        logging.debug("Regular product request failed for %s: %s", url, regular_error)
        log_ajax_fallback_once()
        source = get_text(
            url,
            config.request_timeout,
            referer=referer,
            ajax=True,
            retry_forbidden=True,
        )
    page_data = extract_page_data(source)
    extractor = ProductHTMLExtractor()
    extractor.feed(source)
    detail_images = []
    for raw in extractor.detail_images:
        image_url = canonical_image_url(urljoin(url, raw), large=False)
        if image_url and not any(marker in image_url.lower() for marker in PROMO_IMAGE_MARKERS):
            detail_images.append(image_url)
    description_source = "\n".join(extractor.bullets) if extractor.bullets else " ".join(extractor.text_parts)
    description = clean_description_text(description_source)
    bullets = [clean_description_text(value)[:700] for value in extractor.bullets if clean_description_text(value)][:6]
    if not description:
        description = clean_description_text(extractor.meta_description)
    if not description:
        description = clean_description_text(str(page_data.get("name") or ""))

    images = [canonical_image_url(page_data.get("image_url", ""))]
    images.extend(canonical_image_url(urljoin(url, raw)) for raw in extractor.main_images)

    raw_variants = page_data.get("variants") or []
    variants: list[dict[str, Any]] = []
    for variant in raw_variants:
        if not variant.get("enable_sell", True) or variant.get("availability") in {"out_of_stock", "unavailable"}:
            continue
        params = {
            normalize_space(str(p.get("param_name", ""))): normalize_space(str(p.get("param_value", "")))
            for p in (variant.get("parameters") or [])
        }
        variants.append(
            {
                "id": str(variant.get("id") or ""),
                "parameters": params,
                "price": float(variant.get("discount_price") or variant.get("price") or page_data.get("discount_price") or page_data.get("price") or 0),
                "list_price": float(variant.get("price") or page_data.get("price") or 0) or None,
            }
        )

    if not variants:
        variants = [
            {
                "id": str(page_data.get("id") or "single"),
                "parameters": {},
                "price": float(page_data.get("discount_price") or page_data.get("price") or 0),
                "list_price": float(page_data.get("price") or 0) or None,
            }
        ]

    base_price = float(page_data.get("discount_price") or page_data.get("price") or 0)
    regular_price = float(page_data.get("price") or 0) or None
    return Product(
        url=url,
        product_id=str(page_data.get("id") or ""),
        name=normalize_space(str(page_data.get("name") or "")),
        category_path=normalize_space(str(page_data.get("category_path") or page_data.get("category") or "")),
        description=description,
        bullets=unique_strings(bullets),
        main_images=unique_strings(image for image in images if image),
        detail_images=unique_strings(detail_images),
        price=base_price,
        list_price=regular_price if regular_price and regular_price > base_price else None,
        variants=variants,
    )


def normalize_size(raw: str) -> tuple[str, str, str, bool]:
    raw = normalize_space(raw)
    compact = raw.upper().replace(" ", "")
    numeric_match = re.fullmatch(r"(\d{2,3})(?:СМ|CM)?", compact)
    if numeric_match:
        number = numeric_match.group(1)
        return "2 - Regular Size", "7 - Numeric", number, True

    age_match = re.fullmatch(r"(\d{1,2})(?:Г|ГОД|Y)", compact)
    if age_match:
        return "2 - Regular Size", "8 - Age", f"{age_match.group(1)}Y", True

    normalized = SIZE_ALIASES.get(compact, compact)
    if normalized in {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"}:
        return "2 - Regular Size", "10 - Alpha", normalized, False
    if normalized in {"ONE SIZE", "ONESIZE", "УНИВЕРСАЛЕН", "ЕДИН РАЗМЕР", ""}:
        return "2 - Regular Size", "1 - One Size", "One Size", False
    return "101 - Custom size", "10 - Alpha", raw or "One Size", False


def get_variant_value(parameters: dict[str, str], names: Iterable[str]) -> str:
    lowered = {key.lower(): value for key, value in parameters.items()}
    for name in names:
        if name.lower() in lowered:
            return lowered[name.lower()]
    return ""


def classify_kind(product: Product) -> str | None:
    text = f"{product.category_path} {product.name} {product.url}".lower()
    if re.search(r"раниц|ranic|backpack", text):
        return None
    if re.search(r"шапк|shapk|cap\b|hat\b", text):
        return "hat"
    if re.search(r"къс(?:и|) екип|кас[- ]?екип|short set", text):
        return "short_set"
    if re.search(r"дълг(?:и|) екип|далг[- ]?екип|long set", text):
        return "long_set"
    if re.search(r"шорт|shorti|shorts", text):
        return "shorts"
    if re.search(r"долнищ|dolnish|пантал", text):
        return None
    if re.search(r"ветров|vetrov|windbreak", text):
        return "windbreaker"
    if re.search(r"суичър без ръкав|suichar-bez-rakavi|елек|elek|\bvest\b", text):
        return "vest"
    if re.search(r"потник|potnik|tank top|tank-top", text):
        return "tank"
    if re.search(r"\bполо\b|\bpolo\b", text):
        return "polo"
    if re.search(r"суич|suich|hoodie", text):
        desc = product.description.lower()
        return "hoodie" if "качул" in desc or "hood" in desc else "sweatshirt"
    if re.search(r"блуз|bluz|sweatshirt", text):
        return "sweatshirt"
    if re.search(r"тениск|tenisk|t-shirt|tshirt|tee\b", text):
        return "tshirt"
    return None


def category_for(kind: str, is_kids: bool) -> tuple[str, str] | None:
    if kind == "hat":
        key = "hat"
    elif kind in {"short_set", "long_set"}:
        key = "set_kids" if is_kids else f"{kind}_adult"
    else:
        key = f"{kind}_{'kids' if is_kids else 'adult'}"
    category_id = SUPPORTED_CATEGORY_IDS.get(key)
    return (category_id, key) if category_id else None


def parse_composition(text: str, kind: str, config: Config) -> dict[str, float]:
    lowered = normalize_space(text).lower()
    found: dict[str, float] = {}
    patterns = (
        r"(\d{1,3}(?:[.,]\d+)?)\s*%\s*([a-zа-я\- ]{3,30})",
        r"([a-zа-я\- ]{3,30})\s*(\d{1,3}(?:[.,]\d+)?)\s*%",
    )
    for pattern_index, pattern in enumerate(patterns):
        for match in re.finditer(pattern, lowered, flags=re.IGNORECASE):
            number_text, material_text = (match.group(1), match.group(2)) if pattern_index == 0 else (match.group(2), match.group(1))
            number = float(number_text.replace(",", "."))
            for needle, standard in MATERIAL_NAMES.items():
                if needle in material_text:
                    found[standard] = number
                    break
    if found and 99.0 <= sum(found.values()) <= 101.0:
        return found
    for needle, standard in MATERIAL_NAMES.items():
        if f"100% {needle}" in lowered or f"100 % {needle}" in lowered:
            return {standard: 100.0}
    if "полиестер" in lowered and "памук" not in lowered:
        return {"Polyester": 100.0}
    return dict(config.default_composition.get(kind, {"Cotton": 100.0}))


def age_group_for(size: str) -> str:
    number_match = re.search(r"\d+", size)
    if number_match and int(number_match.group()) >= 158:
        return "13 years and above"
    return "12 and under"


def product_to_rows(product: Product, config: Config) -> tuple[list[OutputRow], str | None]:
    kind = classify_kind(product)
    if not kind:
        return [], "No matching product category is available in the supplied Temu template"

    rows: list[OutputRow] = []
    composition = parse_composition(f"{product.name}\n{product.description}", kind, config)
    for variant in product.variants:
        params = variant.get("parameters") or {}
        raw_size = get_variant_value(params, ("Размер", "Size", "Възраст", "Age"))
        family, sub_family, temu_size, size_implies_kids = normalize_size(raw_size)
        path_implies_kids = "ЗА ДЕЦА" in product.category_path.upper() or "ДЕТСК" in product.name.upper()
        is_kids = size_implies_kids or path_implies_kids
        category = category_for(kind, is_kids)
        if not category:
            continue
        category_id, category_key = category
        variant_id = str(variant.get("id") or "single")
        parent_sku = f"TS-{product.product_id}-{category_id}"
        sku = f"{parent_sku}-{variant_id}"
        display_name = product.name
        if any(normalize_size(get_variant_value(v.get("parameters") or {}, ("Размер", "Size")))[3] != is_kids for v in product.variants):
            display_name += " - за деца" if is_kids else " - за възрастни"
        color = get_variant_value(params, ("Цвят", "Color")) or config.default_color
        price = float(variant.get("price") or product.price)
        regular = variant.get("list_price") or product.list_price
        list_price = float(regular) if regular and float(regular) > price else None
        rows.append(
            OutputRow(
                source_url=product.url,
                source_category=product.category_path,
                category_id=category_id,
                category_key=category_key,
                product_id=product.product_id,
                parent_sku=parent_sku,
                sku=sku,
                name=display_name[:500],
                description=product.description,
                bullets=product.bullets,
                main_images=product.main_images,
                detail_images=product.detail_images,
                size_raw=raw_size,
                size_family=family,
                sub_size_family=sub_family,
                temu_size=temu_size,
                color=color,
                price=price,
                list_price=list_price,
                quantity=config.default_quantity,
                composition=composition,
                product_kind=kind,
                is_kids=is_kids,
                age_group=age_group_for(temu_size),
            )
        )
    return rows, None if rows else "No sellable variant could be mapped"


class TemuTemplate:
    def __init__(self, workbook: Any):
        self.workbook = workbook
        self.sheet = workbook[TEMPLATE_SHEET]
        self.mode_sheet = workbook["GoodsLevelMode"]
        self.category_names = {
            str(row[0].value): str(row[1].value)
            for row in workbook["Category Name"].iter_rows(min_row=1, max_col=2)
            if row[0].value is not None and row[1].value is not None
        }
        self.mode_rows = {
            str(self.mode_sheet.cell(row, 1).value): row
            for row in range(2, self.mode_sheet.max_row + 1)
            if self.mode_sheet.cell(row, 1).value is not None
        }
        self.required_cache: dict[str, list[int]] = {}
        self.headers = {col: str(self.sheet.cell(2, col).value or "") for col in range(1, self.sheet.max_column + 1)}

    def required_columns(self, category_id: str) -> list[int]:
        if category_id in self.required_cache:
            return self.required_cache[category_id]
        mode_row = self.mode_rows.get(f"{category_id}_require")
        if not mode_row:
            raise ValueError(f"Category {category_id} is not included in this Temu template")
        required = [
            col
            for col in range(1, self.sheet.max_column + 1)
            if self.mode_sheet.cell(mode_row, col).value == "require"
        ]
        self.required_cache[category_id] = required
        return required

    def write_row(self, excel_row: int, item: OutputRow, config: Config) -> None:
        ws = self.sheet
        values: dict[int, Any] = {
            5: item.category_id,
            12: item.name,
            13: item.parent_sku,
            14: item.sku,
            20: item.description,
            column_index_from_string("AWR"): "Color" if item.product_kind == "hat" else "Color × Size",
            column_index_from_string("AWV"): item.color,
            column_index_from_string("BKW"): item.quantity,
            column_index_from_string("BKX"): round(item.price, 2),
            column_index_from_string("BKY"): item.source_url,
            column_index_from_string("BLG"): "Yes",
            column_index_from_string("BLH"): 1,
            column_index_from_string("BLI"): "piece",
            column_index_from_string("BLQ"): config.shipping_template,
            column_index_from_string("BLU"): config.country_of_origin,
            column_index_from_string("BNU"): item.sku,
            column_index_from_string("BNV"): config.manufacturer,
        }
        if config.eu_responsible_person:
            values[column_index_from_string("BNW")] = config.eu_responsible_person
        if item.product_kind != "hat":
            values.update(
                {
                    column_index_from_string("AWS"): item.size_family,
                    column_index_from_string("AWT"): item.sub_size_family,
                    column_index_from_string("AWU"): item.temu_size,
                    column_index_from_string("AXL"): "cm-g-ml",
                }
            )
        if item.list_price and item.list_price > item.price:
            values[column_index_from_string("BKZ")] = round(item.list_price, 2)
        else:
            values[column_index_from_string("BLA")] = "N/A"

        package = DEFAULT_PACKAGE.get(item.product_kind, (500.0, 35.0, 30.0, 7.0))
        values[column_index_from_string("BLB")] = package[0]
        values[column_index_from_string("BLC")] = package[1]
        values[column_index_from_string("BLD")] = package[2]
        values[column_index_from_string("BLE")] = package[3]

        for offset, bullet in enumerate(item.bullets[:6]):
            values[21 + offset] = bullet[:700]
        for offset, image_url in enumerate(item.detail_images[:95]):
            values[27 + offset] = image_url
        first_sku_image = column_index_from_string("BKL")
        for offset, image_url in enumerate(item.main_images[:10]):
            values[first_sku_image + offset] = image_url

        required = self.required_columns(item.category_id)
        material = max(item.composition, key=item.composition.get) if item.composition else "Cotton"
        for col in required:
            title = self.headers.get(col, "")
            lowered = title.lower()
            if title == "12 - Material":
                values[col] = material
            elif "composition:" in lowered or "ingredient:" in lowered or "ingredients:" in lowered:
                component = title.rsplit(":", 1)[-1].strip()
                values[col] = round(float(item.composition.get(component, 0.0)), 1)
            elif "applicable age group" in lowered:
                values[col] = item.age_group
            elif "size chart element" in str(ws.cell(4, col).value or "").lower() or " - product - " in lowered:
                values[col] = self.measurement_value(title, item)

        for col, value in values.items():
            if value not in (None, ""):
                ws.cell(excel_row, col).value = value

    @staticmethod
    def measurement_value(title: str, item: OutputRow) -> float:
        size = item.temu_size
        adult_size = SIZE_ALIASES.get(size.upper(), size.upper())
        if item.is_kids:
            top_table = (
                TOP_MEASUREMENTS_HOODIE_KIDS
                if item.product_kind in {"hoodie", "sweatshirt"}
                else TOP_MEASUREMENTS_TSHIRT_KIDS
            )
            chest, top_length = top_table.get(size, (43.0, 55.0))
            numeric = int(re.search(r"\d+", size).group()) if re.search(r"\d+", size) else 128
            if item.product_kind in {"short_set", "long_set", "set"}:
                waist, pants_length = BOTTOM_MEASUREMENTS_SET_KIDS.get(size, (26.0, 32.0))
            else:
                waist = 23.0 + max(0, numeric - 104) * 0.125
                pants_length = 28.0 + max(0, numeric - 104) * 0.18
            hip = waist + 14.0
            inseam = max(1.0, pants_length - 20.0)
        else:
            source_text = f"{item.source_category} {item.name}".lower()
            if item.product_kind in {"hoodie", "sweatshirt"}:
                top_table = TOP_MEASUREMENTS_HOODIE_ADULT
            elif item.product_kind == "tank":
                top_table = TOP_MEASUREMENTS_TANK_ADULT
            elif item.product_kind == "windbreaker":
                top_table = TOP_MEASUREMENTS_WINDBREAKER_ADULT
            elif item.product_kind == "vest":
                top_table = TOP_MEASUREMENTS_VEST_ADULT
            elif "дам" in source_text:
                top_table = TOP_MEASUREMENTS_TSHIRT_WOMEN
            else:
                top_table = TOP_MEASUREMENTS_TSHIRT_ADULT
            chest, top_length = top_table.get(adult_size, (54.0, 73.0))
            index = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"].index(adult_size) if adult_size in {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL"} else 4
            if item.product_kind == "shorts":
                waist, pants_length = BOTTOM_MEASUREMENTS_SHORTS_ADULT.get(adult_size, (42.5, 48.0))
            elif item.product_kind in {"short_set", "set"}:
                waist, pants_length = BOTTOM_MEASUREMENTS_SET_ADULT.get(adult_size, (36.0, 48.0))
            else:
                waist = 30.0 + index * 2.0
                pants_length = 98.0 + index * 2.0
            hip = 44.0 + index * 2.5
            inseam = max(1.0, pants_length - 25.0) if item.product_kind in {"shorts", "short_set", "set"} else 72.0 + index
        lowered = title.lower()
        if "chest" in lowered:
            return round(chest, 1)
        if "waist" in lowered:
            return round(waist, 1)
        if "hip" in lowered:
            return round(hip, 1)
        if "inseam" in lowered:
            return round(inseam, 1)
        if "pants length" in lowered or "underpants" in lowered:
            return round(pants_length, 1)
        if "foot length" in lowered:
            return 24.0
        if "socks" in lowered and "width" in lowered:
            return 8.0
        if "socks" in lowered:
            return 20.0
        if "length" in lowered:
            return round(top_length, 1)
        if "width" in lowered:
            return round(chest, 1)
        return 1.0


def chunks(values: list[OutputRow], size: int) -> Iterable[list[OutputRow]]:
    for start in range(0, len(values), size):
        yield values[start : start + size]


def write_raw_export(rows: list[OutputRow], path: Path) -> None:
    fields = [
        "source_url", "source_category", "category_id", "category_key", "product_id",
        "parent_sku", "sku", "name", "size_raw", "temu_size", "color", "price",
        "list_price", "quantity", "composition", "main_images", "detail_images", "description",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for item in rows:
            writer.writerow(
                {
                    "source_url": item.source_url,
                    "source_category": item.source_category,
                    "category_id": item.category_id,
                    "category_key": item.category_key,
                    "product_id": item.product_id,
                    "parent_sku": item.parent_sku,
                    "sku": item.sku,
                    "name": item.name,
                    "size_raw": item.size_raw,
                    "temu_size": item.temu_size,
                    "color": item.color,
                    "price": item.price,
                    "list_price": item.list_price or "",
                    "quantity": item.quantity,
                    "composition": json.dumps(item.composition, ensure_ascii=False),
                    "main_images": " | ".join(item.main_images),
                    "detail_images": " | ".join(item.detail_images),
                    "description": item.description,
                }
            )


def write_skipped(rows: list[dict[str, str]], path: Path) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["url", "reason"])
        writer.writeheader()
        writer.writerows(rows)


def write_workbooks(rows: list[OutputRow], template_path: Path, output_dir: Path, config: Config) -> list[Path]:
    created: list[Path] = []
    for part_number, batch in enumerate(chunks(rows, config.max_rows_per_file), start=1):
        destination = output_dir / f"TEMU_TOPSTOKEE_UPLOAD_part_{part_number:03d}.xlsx"
        shutil.copy2(template_path, destination)
        workbook = load_workbook(destination, read_only=False, data_only=False)
        helper = TemuTemplate(workbook)
        for offset, item in enumerate(batch):
            helper.write_row(FIRST_DATA_ROW + offset, item, config)
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(destination)
        workbook.close()
        created.append(destination)
        logging.info("Created %s with %d rows", destination.name, len(batch))
    return created


def configure_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    handlers = [logging.StreamHandler(sys.stdout), logging.FileHandler(output_dir / "scraper.log", encoding="utf-8")]
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--template", type=Path, default=Path("template.xlsx"))
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--output-dir", type=Path, default=Path("outputs"))
    parser.add_argument("--mode", choices=("test", "full"), default="test")
    parser.add_argument("--max-products", type=int, default=0, help="0 means all products in full mode")
    parser.add_argument("--workers", type=int, default=0)
    parser.add_argument("--start-url", action="append", default=[], help="Optional product URL; repeat as needed")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = Config.from_json(args.config)
    requested_workers = args.workers or config.workers
    if args.workers > 0:
        config.workers = max(1, min(args.workers, 3))
    configure_logging(args.output_dir)
    if requested_workers != config.workers:
        logging.warning(
            "Workers reduced from %d to %d to avoid CloudCart rate blocking",
            requested_workers,
            config.workers,
        )
    if not args.template.exists():
        logging.error("Template file not found: %s", args.template)
        return 2

    limit = args.max_products or (15 if args.mode == "test" else 0)
    urls = list(dict.fromkeys(args.start_url)) if args.start_url else discover_product_urls(config, limit)
    if limit:
        urls = urls[:limit]
    logging.info("Products selected: %d", len(urls))

    products: list[Product] = []
    skipped: list[dict[str, str]] = []
    started = time.time()
    with ThreadPoolExecutor(max_workers=config.workers) as executor:
        future_to_url = {executor.submit(parse_product, url, config): url for url in urls}
        for index, future in enumerate(as_completed(future_to_url), start=1):
            url = future_to_url[future]
            try:
                product = future.result()
                products.append(product)
                logging.info("Parsed %d/%d: %s", index, len(urls), product.name)
            except Exception as exc:
                logging.exception("Failed: %s", url)
                skipped.append({"url": url, "reason": f"Request/parse error: {exc}"})
            if config.request_delay_seconds:
                time.sleep(config.request_delay_seconds)

    output_rows: list[OutputRow] = []
    for product in sorted(products, key=lambda item: item.url):
        mapped, reason = product_to_rows(product, config)
        output_rows.extend(mapped)
        if reason:
            skipped.append({"url": product.url, "reason": reason})

    output_rows.sort(key=lambda item: (item.parent_sku, item.sku))
    write_raw_export(output_rows, args.output_dir / "topstokee_raw_export.csv")
    write_skipped(skipped, args.output_dir / "topstokee_skipped_products.csv")
    created = write_workbooks(output_rows, args.template, args.output_dir, config) if output_rows else []

    summary = {
        "products_requested": len(urls),
        "products_parsed": len(products),
        "temu_rows": len(output_rows),
        "skipped_products": len(skipped),
        "xlsx_parts": len(created),
        "elapsed_seconds": round(time.time() - started, 1),
    }
    (args.output_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    logging.info("Summary: %s", json.dumps(summary, ensure_ascii=False))
    return 0 if products else 1


if __name__ == "__main__":
    raise SystemExit(main())
