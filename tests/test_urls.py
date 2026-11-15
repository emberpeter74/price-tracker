import pytest

from pricetracker.urls import clean_link, host_of, is_product_link, same_site


def test_clean_link_drops_tracking_keeps_variant():
    u = ("https://thetankshop.co.uk/products/harlequin-650?variant=37702586237120"
         "&currency=GBP&utm_source=google&utm_medium=cpc&gad_source=4&gclid=abc&srsltid=x")
    assert clean_link(u) == "https://thetankshop.co.uk/products/harlequin-650?variant=37702586237120"


def test_clean_link_keeps_attribute_selectors():
    u = "https://www.landownerproducts.co.uk/adblue-1000l?attr%5B146%5D=4&ps=MT"
    assert "attr%5B146%5D=4" in clean_link(u) and "ps=" not in clean_link(u)


def test_clean_link_drops_fragment_and_search_param():
    assert clean_link("https://a.co.uk/p/x?search=piusi%20k33#reviews") == "https://a.co.uk/p/x"


def test_clean_link_leaves_non_urls_alone():
    assert clean_link("#N/A") == "#N/A"
    assert clean_link("  https://a.co.uk/p/x  ") == "https://a.co.uk/p/x"


def test_host_of():
    assert host_of("https://www.FuelTankShop.co.uk/x") == "fueltankshop.co.uk"
    assert host_of("fueltankshop.co.uk") == "" and host_of("#N/A") == ""


def test_same_site_ignores_www_and_path():
    assert same_site("https://www.manutan.co.uk/en/key", "https://manutan.co.uk/p/1")
    assert not same_site("https://a.co.uk/", "https://b.co.uk/")
    assert not same_site("", "")


@pytest.mark.parametrize("u,ok", [
    ("https://www.chemfastuk.com/products/air1-adblue", True),
    ("https://www.fueltankshop.co.uk/1000-litre-waste-oil-tank-harlequin-orb1000/p24", True),
    ("https://www.agriemach.com/553000-piusi-k33-fuel-flow-meter", True),
    ("https://www.chemfastuk.com/collections/fuelbox", False),
    ("https://www.plumbnation.co.uk/oil-tanks/", False),
    ("https://www.fueltankshop.co.uk/", False),
    ("petroy.com", False), ("#N/A", False), ("", False),
])
def test_is_product_link(u, ok):
    assert is_product_link(u) is ok
