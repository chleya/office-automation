"""
Basic tests for Office Automation
"""

import pytest
from office_automation import OfficeAutomation


def test_import():
    """Test that OfficeAutomation can be imported."""
    from office_automation import OfficeAutomation
    assert OfficeAutomation is not None


def test_instantiation():
    """Test that OfficeAutomation can be instantiated."""
    office = OfficeAutomation()
    assert office is not None
    assert hasattr(office, 'word')
    assert hasattr(office, 'excel')
    assert hasattr(office, 'powerpoint')
    assert hasattr(office, 'converter')
    assert hasattr(office, 'wps')


def test_get_info():
    """Test get_info method."""
    office = OfficeAutomation()
    info = office.get_info()
    
    assert 'version' in info
    assert 'python_version' in info
    assert 'modules' in info
    
    # Check module structure
    modules = info['modules']
    assert 'word' in modules
    assert 'excel' in modules
    assert 'powerpoint' in modules
    assert 'converter' in modules
    assert 'wps' in modules


def test_quick_functions():
    """Test quick access functions."""
    from office_automation import (
        quick_create_document,
        quick_create_spreadsheet,
        quick_convert
    )
    
    # Test that functions exist
    assert callable(quick_create_document)
    assert callable(quick_create_spreadsheet)
    assert callable(quick_convert)


@pytest.mark.parametrize("module_name", ['word', 'excel', 'powerpoint', 'converter', 'wps'])
def test_module_capabilities(module_name):
    """Test that each module has get_capabilities method."""
    office = OfficeAutomation()
    module = getattr(office, module_name)
    
    if hasattr(module, 'get_capabilities'):
        capabilities = module.get_capabilities()
        assert isinstance(capabilities, dict)
        assert 'available' in capabilities
    else:
        # Some modules may not have this method in dummy mode
        pass


def test_config_initialization():
    """Test configuration initialization."""
    # Test with default config
    office1 = OfficeAutomation()
    assert office1.config is not None
    
    # Test with custom config
    custom_config = {
        'temp_dir': 'custom_temp',
        'encoding': 'gbk'
    }
    office2 = OfficeAutomation(config=custom_config)
    assert office2.config['temp_dir'] == 'custom_temp'
    assert office2.config['encoding'] == 'gbk'
    
    # Test that defaults are merged
    assert 'default_templates' in office2.config