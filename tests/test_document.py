import pytest
from documents import Document

def test_document():
    Document()
    value = 1
    assert 1 == value