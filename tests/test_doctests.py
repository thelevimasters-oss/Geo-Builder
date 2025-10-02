import doctest

import deed_extractor


def test_deed_extractor_doctests():
    failure_count, _ = doctest.testmod(deed_extractor, optionflags=doctest.ELLIPSIS)
    assert failure_count == 0
