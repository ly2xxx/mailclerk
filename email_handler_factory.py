import importlib
import logging
from typing import Dict, List, Optional, Type
from email_handler_base import EmailHandlerBase

logger = logging.getLogger(__name__)

class EmailHandlerFactory:
    """Factory class for dynamically discovering and creating email handlers"""
    
    def __init__(self):
        self._handlers = {}
        self._discover_handlers()
    
    def _discover_handlers(self):
        """Dynamically discover available email handlers"""
        self._handlers = {}
        
        # List of potential handler modules to check
        potential_handlers = [
            {
                'key': 'browser',
                'module': 'email_browser_handler',
                'class': 'BrowserEmailHandler',
                'name': '🌐 Browser/Web (Outlook.com) - Recommended',
                'description': 'Uses Microsoft Graph API\n• Works on any platform\n• No Outlook installation needed\n• Requires web authentication',
                'platform_requirements': ['any'],
                'dependencies': ['requests']
            },
            {
                'key': 'desktop',
                'module': 'email_handler',
                'class': 'OutlookEmailHandler',
                'name': '💻 Desktop (Outlook COM) - Windows Only',
                'description': 'Requires Windows + Outlook\n• Local Outlook installation\n• COM registration needed\n• May need Administrator rights',
                'platform_requirements': ['Windows'],
                'dependencies': ['win32com.client']
            }
        ]
        
        for handler_info in potential_handlers:
            try:
                # Check platform requirements
                if not self._check_platform_compatibility(handler_info['platform_requirements']):
                    logger.info(f"Handler {handler_info['key']} not compatible with current platform")
                    continue
                
                # Check dependencies
                if not self._check_dependencies(handler_info['dependencies']):
                    logger.info(f"Handler {handler_info['key']} missing required dependencies")
                    continue
                
                # Try to import the module
                module = importlib.import_module(handler_info['module'])
                handler_class = getattr(module, handler_info['class'])
                
                # Verify it's a subclass of EmailHandlerBase
                if issubclass(handler_class, EmailHandlerBase):
                    self._handlers[handler_info['key']] = {
                        'class': handler_class,
                        'name': handler_info['name'],
                        'description': handler_info['description'],
                        'module': handler_info['module'],
                        'class_name': handler_info['class']
                    }
                    logger.info(f"Discovered email handler: {handler_info['key']}")
                else:
                    logger.warning(f"Handler {handler_info['class']} does not inherit from EmailHandlerBase")
                    
            except ImportError as e:
                logger.info(f"Handler {handler_info['key']} not available (module import failed): {e}")
            except AttributeError as e:
                logger.warning(f"Handler {handler_info['key']} class not found: {e}")
            except Exception as e:
                logger.error(f"Error discovering handler {handler_info['key']}: {e}")
    
    def _check_platform_compatibility(self, requirements: List[str]) -> bool:
        """Check if current platform meets handler requirements"""
        if 'any' in requirements:
            return True
        
        import platform
        current_platform = platform.system()
        
        return current_platform in requirements
    
    def _check_dependencies(self, dependencies: List[str]) -> bool:
        """Check if required dependencies are available"""
        for dependency in dependencies:
            try:
                importlib.import_module(dependency)
            except ImportError:
                return False
        return True
    
    def get_available_handlers(self) -> Dict[str, Dict]:
        """Get dictionary of available handlers"""
        return self._handlers.copy()
    
    def get_handler_names(self) -> Dict[str, str]:
        """Get mapping of handler keys to display names"""
        return {key: info['name'] for key, info in self._handlers.items()}
    
    def get_handler_keys(self) -> List[str]:
        """Get list of available handler keys"""
        return list(self._handlers.keys())
    
    def get_handler_info(self, handler_key: str) -> Optional[Dict]:
        """Get information about a specific handler"""
        return self._handlers.get(handler_key)
    
    def create_handler(self, handler_key: str, **kwargs) -> Optional[EmailHandlerBase]:
        """Create an instance of the specified handler"""
        if handler_key not in self._handlers:
            logger.error(f"Handler '{handler_key}' not available")
            return None
        
        try:
            handler_class = self._handlers[handler_key]['class']
            return handler_class(**kwargs)
        except Exception as e:
            logger.error(f"Failed to create handler '{handler_key}': {e}")
            return None
    
    def is_handler_available(self, handler_key: str) -> bool:
        """Check if a specific handler is available"""
        return handler_key in self._handlers
    
    def get_default_handler_key(self) -> Optional[str]:
        """Get the default handler key (prefer browser, fallback to first available)"""
        if 'browser' in self._handlers:
            return 'browser'
        elif self._handlers:
            return list(self._handlers.keys())[0]
        return None
    
    def refresh_handlers(self):
        """Re-discover available handlers (useful if modules are added/removed)"""
        logger.info("Refreshing handler discovery...")
        self._discover_handlers()
        logger.info(f"Found {len(self._handlers)} available handlers: {list(self._handlers.keys())}")

# Global factory instance
_handler_factory = None

def get_handler_factory() -> EmailHandlerFactory:
    """Get the global handler factory instance"""
    global _handler_factory
    if _handler_factory is None:
        _handler_factory = EmailHandlerFactory()
    return _handler_factory

def create_email_handler(handler_key: str, **kwargs) -> Optional[EmailHandlerBase]:
    """Convenience function to create an email handler"""
    factory = get_handler_factory()
    return factory.create_handler(handler_key, **kwargs)

def get_available_handlers() -> Dict[str, Dict]:
    """Convenience function to get available handlers"""
    factory = get_handler_factory()
    return factory.get_available_handlers()

def get_handler_display_names() -> Dict[str, str]:
    """Convenience function to get handler display names"""
    factory = get_handler_factory()
    return factory.get_handler_names()