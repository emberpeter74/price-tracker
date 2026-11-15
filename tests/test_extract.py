from decimal import Decimal as D

from pricetracker.extract import parse_price, parse_product_page

URL = "https://shop.example/p/1"


def ld(obj_json):
    return f'<html><head><script type="application/ld+json">{obj_json}</script></head><body><h1>Thing</h1></body></html>'


# ----- real pages -----

def test_agriemach_jsonld_gives_price_and_mpn(load_page):
    p = parse_product_page(load_page("agriemach_k33.html"), URL)
    assert p.is_single_product and p.mpn == "553000" and p.name == "Piusi K33 Fuel Flow Meter"
    assert [o.price for o in p.offers] == [D("146.14")] and p.offers[0].currency == "GBP"
    assert D("175.37") in p.text_prices and "553000" in p.text


def test_hytek_gives_sku_but_no_readable_price(load_page):
    p = parse_product_page(load_page("hytek_k33.html"), URL)
    assert p.is_single_product and p.sku == "K33" and p.mpn == "" and p.offers == []
    assert "&amp;" not in p.name and "Piusi K33" in p.name


def test_fueltankshop_has_no_jsonld_product_and_falls_back_to_meta(load_page):
    p = parse_product_page(load_page("fueltankshop_orb1000.html"), URL)
    assert p.is_single_product and [o.price for o in p.offers] == [D("1390.00")]
    assert "ORB1000" in p.name and p.mpn == "" and p.sku == ""


def test_thetankshop_gives_sku_brand_and_price(load_page):
    p = parse_product_page(load_page("thetankshop_1000ite.html"), URL)
    assert p.sku == "1000ITE" and p.brand == "Harlequin"
    assert [o.price for o in p.offers] == [D("1820.40")] and D("1517.00") in p.text_prices


def test_chemfast_offer_list(load_page):
    p = parse_product_page(load_page("chemfast_adblue.html"), URL)
    assert [o.price for o in p.offers] == [D("399.6")] and D("333") in p.text_prices


def test_tanksforeverything_empty_codes_are_empty_strings(load_page):
    p = parse_product_page(load_page("tanksforeverything_exocet.html"), URL)
    assert p.is_single_product and p.mpn == "" and p.gtin == "" and p.offers[0].price == D("52.74")


# ----- shapes -----

def test_multi_offer_page_returns_every_offer_with_its_codes():
    p = parse_product_page(ld('''{"@type":"Product","name":"Camlock","offers":[
        {"@type":"Offer","price":"2.70","priceCurrency":"GBP","sku":"F100A","url":"/p/1?variant=11"},
        {"@type":"Offer","price":"4.10","priceCurrency":"GBP","sku":"F150A","mpn":"M150","gtin13":"5056345600159"}]}'''), URL)
    assert [(o.price, o.sku, o.mpn, o.gtin) for o in p.offers] == [
        (D("2.70"), "F100A", "", ""), (D("4.10"), "F150A", "M150", "5056345600159")]
    assert p.offers[0].url == "https://shop.example/p/1?variant=11"


def test_jsonld_graph_and_list_wrappers_are_unwrapped():
    a = parse_product_page(ld('{"@graph":[{"@type":"WebSite"},{"@type":["Product","Thing"],"name":"X","sku":"S1","offers":{"price":10}}]}'), URL)
    b = parse_product_page(ld('[{"@type":"BreadcrumbList"},{"@type":"Product","name":"X","sku":"S1","offers":{"price":"10.00"}}]'), URL)
    c = parse_product_page(ld('{"@type":"WebPage","mainEntity":{"@type":"Product","name":"X","sku":"S1","offers":{"price":"10"}}}'), URL)
    for p in (a, b, c):
        assert p.is_single_product and p.sku == "S1" and p.offers[0].price == D("10")


def test_product_group_variants_each_contribute_offers():
    p = parse_product_page(ld('''{"@type":"ProductGroup","name":"Hose","hasVariant":[
        {"@type":"Product","sku":"H25","offers":{"price":"32.00","priceCurrency":"GBP"}},
        {"@type":"Product","sku":"H32","mpn":"VAC/32","offers":{"price":"41.00","priceCurrency":"GBP"}}]}'''), URL)
    assert p.is_single_product
    assert [(o.sku, o.mpn, o.price) for o in p.offers] == [("H25", "", D("32.00")), ("H32", "VAC/32", D("41.00"))]


def test_aggregate_offer_yields_low_and_high_as_two_offers():
    p = parse_product_page(ld('{"@type":"Product","name":"X","offers":{"@type":"AggregateOffer","lowPrice":"10.00","highPrice":"15.00","priceCurrency":"GBP"}}'), URL)
    assert [o.price for o in p.offers] == [D("10.00"), D("15.00")]
    q = parse_product_page(ld('{"@type":"Product","name":"X","offers":{"@type":"AggregateOffer","lowPrice":"10.00","highPrice":"10.00"}}'), URL)
    assert [o.price for o in q.offers] == [D("10.00")] and q.offers[0].currency == "GBP"


def test_price_specification_is_read():
    p = parse_product_page(ld('{"@type":"Product","name":"X","offers":{"priceSpecification":{"price":"99.50","priceCurrency":"GBP"}}}'), URL)
    assert p.offers[0].price == D("99.50")


def test_price_strings_with_pound_sign_and_commas_parse():
    assert parse_price("£1,390.00") == D("1390.00") and parse_price(" 52.74 ") == D("52.74")
    assert parse_price(399.6) == D("399.6") and parse_price("1 234,50") is None
    assert parse_price("") is None and parse_price(None) is None and parse_price("POA") is None
    assert parse_price("0") is None and parse_price("-5") is None
    assert parse_price("0.01") is None and parse_price("0.02") == D("0.02")     # 0.01 = "call for price"


def test_microdata_fallback():
    html = '''<html><body><div itemscope itemtype="https://schema.org/Product">
      <h1 itemprop="name">Cim-Tek 70062</h1><span itemprop="mpn">70062</span>
      <span itemprop="gtin13" content="0662256700623"></span>
      <div itemprop="offers" itemscope><meta itemprop="priceCurrency" content="GBP">
      <span itemprop="price" content="33.00">£33.00</span></div></div></body></html>'''
    p = parse_product_page(html, URL)
    assert p.is_single_product and p.mpn == "70062" and p.gtin == "0662256700623"
    assert p.name == "Cim-Tek 70062" and [o.price for o in p.offers] == [D("33.00")]


def test_meta_tag_fallback():
    html = '''<html><head><meta property="og:type" content="product">
      <meta property="og:title" content="Harlequin 1400FS">
      <meta property="product:price:amount" content="1,825.00">
      <meta property="product:price:currency" content="GBP"></head><body></body></html>'''
    p = parse_product_page(html, URL)
    assert p.is_single_product and p.name == "Harlequin 1400FS" and p.offers[0].price == D("1825.00")


def test_usd_price_keeps_its_currency():
    p = parse_product_page(ld('{"@type":"Product","name":"X","offers":{"price":"10","priceCurrency":"usd"}}'), URL)
    assert p.offers[0].currency == "USD"


def test_listing_page_is_not_single_product():
    two = parse_product_page(ld('[{"@type":"Product","name":"A","offers":{"price":"1"}},{"@type":"Product","name":"B","offers":{"price":"2"}}]'), URL)
    lst = parse_product_page(ld('{"@type":"ItemList","itemListElement":[{"@type":"ListItem","item":{"@type":"Product","name":"A","offers":{"price":"1"}}}]}'), URL)
    plain = parse_product_page("<html><body><h1>Search results</h1><p>£10.00</p></body></html>", URL)
    assert not two.is_single_product and not lst.is_single_product and not plain.is_single_product
    assert plain.offers == []


def test_page_without_price_has_no_offers():
    p = parse_product_page(ld('{"@type":"Product","name":"X","sku":"S","offers":[[]]}'), URL)
    assert p.is_single_product and p.offers == []


def test_malformed_jsonld_is_skipped_not_raised():
    html = ('<script type="application/ld+json">{not json</script>'
            '<script type="application/ld+json">{"@type":"Product","name":"X","offers":{"price":"5"}}</script>')
    assert parse_product_page(html, URL).offers[0].price == D("5")
    assert parse_product_page("", URL).offers == []
    assert parse_product_page("<<<>>> \x00 garbage", URL).is_single_product is False


def test_visible_text_excludes_scripts_and_collects_prices():
    html = '''<html><body><script>var price = "£999.99";</script><style>.a{}</style>
      <p>Now  £1,517.00 exc. VAT</p><p>( £1,820.40 inc VAT )</p></body></html>'''
    p = parse_product_page(html, URL)
    assert p.text == "Now £1,517.00 exc. VAT ( £1,820.40 inc VAT )"
    assert p.text_prices == frozenset({D("1517.00"), D("1820.40")})
