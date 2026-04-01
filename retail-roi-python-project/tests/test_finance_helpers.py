from retail_roi_model.engine import safe_num, npv, irr


def test_safe_num_handles_none():
    assert safe_num(None) == 0.0


def test_npv_returns_float():
    value = npv(0.1, [100, 100, 100])
    assert isinstance(value, float)


def test_irr_returns_number_or_none():
    value = irr([-100, 60, 60, 60])
    assert value is None or isinstance(value, float)
