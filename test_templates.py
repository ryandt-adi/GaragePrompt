#!/usr/bin/env python3
"""Quick test to verify templates.py loads correctly."""

print("Testing templates.py import...")

try:
    from templates import template_registry, PromptTemplate
    print("✓ Successfully imported template_registry and PromptTemplate")
    
    # Test the registry
    print(f"✓ Registry contains {len(template_registry.list_ids())} templates:")
    for tid in template_registry.list_ids():
        template = template_registry.get(tid)
        print(f"   - {tid}: {template.name}")
    
    print("\n✓ All tests passed! templates.py is working correctly.")
    
except ImportError as e:
    print(f"✗ Import error: {e}")
except Exception as e:
    print(f"✗ Error: {e}")
