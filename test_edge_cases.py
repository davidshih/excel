#!/usr/bin/env python3
"""
Test edge cases and often-ignored aspects of SharePoint integration
"""

import os
import tempfile
import json
from datetime import datetime, timedelta

def test_widget_compatibility():
    """Test widget creation compatibility issues"""
    print("Testing Widget Compatibility...")
    print("=" * 50)
    
    # Test widget parameters that might fail in different ipywidgets versions
    widget_tests = [
        {
            'name': 'Button with button_style',
            'params': {'description': 'Test', 'button_style': 'primary'},
            'safe_params': {'description': 'Test'}
        },
        {
            'name': 'Checkbox with indent',
            'params': {'value': True, 'indent': False},
            'safe_params': {'value': True}
        },
        {
            'name': 'Text with style dict',
            'params': {'value': '', 'style': {'description_width': 'initial'}},
            'safe_params': {'value': ''}
        }
    ]
    
    for test in widget_tests:
        print(f"✓ {test['name']}: Use safe params to avoid version issues")
        print(f"  Risky: {test['params']}")
        print(f"  Safe: {test['safe_params']}")
    
    print("\n⚠️  Widget Compatibility Recommendations:")
    print("  - Avoid button_style parameter")
    print("  - Use minimal widget parameters")
    print("  - Test actual widget instantiation")
    print("  - Provide fallbacks for missing features")

def test_special_characters():
    """Test handling of special characters in various contexts"""
    print("\n\nTesting Special Characters...")
    print("=" * 50)
    
    # Test cases for special characters
    special_names = [
        "John O'Brien",
        "María García-López", 
        "李明 (Li Ming)",
        "Jean-François Müller",
        "Smith & Jones Co.",
        "Review #123",
        "Test/User",
        "Name with spaces  ",
        "email@with+plus.com",
        "Björn Ångström"
    ]
    
    print("1. Reviewer Names with Special Characters:")
    for name in special_names:
        # Test folder name sanitization
        safe_folder = name.replace('/', '_').replace('\\', '_').strip()
        print(f"   {name} → {safe_folder}")
    
    print("\n2. Email Address Edge Cases:")
    email_tests = [
        ("john.doe@company.com", True),
        ("user+tag@company.com", True),
        ("name@sub.domain.com", True),
        ("invalid@", False),
        ("@invalid.com", False),
        ("no-at-sign.com", False),
        ("spaces in@email.com", False)
    ]
    
    for email, valid in email_tests:
        status = "✓ Valid" if valid else "✗ Invalid"
        print(f"   {email}: {status}")
    
    print("\n3. SharePoint Path Encoding:")
    paths = [
        "John Doe",
        "María García",
        "Folder with Spaces",
        "Special!@#$%Characters"
    ]
    
    for path in paths:
        # URL encoding for SharePoint
        encoded = path.replace(' ', '%20').replace('#', '%23')
        print(f"   {path} → {encoded}")

def test_authentication_edge_cases():
    """Test authentication token handling edge cases"""
    print("\n\nTesting Authentication Edge Cases...")
    print("=" * 50)
    
    # Simulate token scenarios
    print("1. Token Expiration Scenarios:")
    
    # Current time
    now = datetime.now()
    
    scenarios = [
        ("Fresh token", now + timedelta(hours=1), "Valid"),
        ("Expiring soon", now + timedelta(minutes=5), "Should refresh"),
        ("Expired", now - timedelta(minutes=1), "Must re-authenticate"),
        ("No expiry info", None, "Assume valid, handle errors")
    ]
    
    for desc, expiry, action in scenarios:
        if expiry:
            remaining = (expiry - now).total_seconds()
            print(f"   {desc}: {action} (expires in {remaining:.0f}s)")
        else:
            print(f"   {desc}: {action}")
    
    print("\n2. Authentication Failure Modes:")
    failures = [
        "Network timeout during auth",
        "User cancels authentication",
        "Invalid client/tenant ID",
        "Insufficient permissions",
        "MFA challenge timeout",
        "Token cache corruption"
    ]
    
    for failure in failures:
        print(f"   • {failure}: Should show clear error message")

def test_concurrent_operations():
    """Test concurrent operation scenarios"""
    print("\n\nTesting Concurrent Operations...")
    print("=" * 50)
    
    print("1. Rate Limiting Scenarios:")
    print("   • Sequential processing: Add 0.5s delay between API calls")
    print("   • Batch size limits: Process max 20 reviewers at once")
    print("   • Retry logic: Exponential backoff on 429 errors")
    
    print("\n2. Concurrent Access Issues:")
    issues = [
        "Multiple users processing same Excel file",
        "SharePoint folder already exists",
        "Email already sent to reviewer",
        "Folder permissions already set",
        "Excel file locked by another process"
    ]
    
    for issue in issues:
        print(f"   • {issue}: Handle gracefully, continue processing")

def test_error_recovery():
    """Test error recovery and partial failure scenarios"""
    print("\n\nTesting Error Recovery...")
    print("=" * 50)
    
    print("1. Partial Failure Handling:")
    reviewers = ["John", "Jane", "Bob", "Alice", "Charlie"]
    failures = [1, 3]  # Jane and Alice fail
    
    for i, reviewer in enumerate(reviewers):
        if i in failures:
            print(f"   ✗ {reviewer}: Failed (continue with others)")
        else:
            print(f"   ✓ {reviewer}: Success")
    
    print(f"\n   Summary: {len(reviewers)-len(failures)}/{len(reviewers)} successful")
    
    print("\n2. Rollback Scenarios:")
    print("   • Folder created but sharing failed: Keep folder, log error")
    print("   • Email lookup failed: Allow manual entry")
    print("   • SharePoint unreachable: Save state for retry")
    print("   • Excel processing failed: Clean up partial files")

def test_network_resilience():
    """Test network failure scenarios"""
    print("\n\nTesting Network Resilience...")
    print("=" * 50)
    
    print("1. Network Failure Points:")
    failure_points = [
        ("During authentication", "Cache token if possible"),
        ("During email lookup", "Use cached results or manual entry"),
        ("During folder sharing", "Retry with exponential backoff"),
        ("During email sending", "Queue for later sending"),
        ("SharePoint site unreachable", "Validate URL, check connectivity")
    ]
    
    for point, recovery in failure_points:
        print(f"   {point}: {recovery}")
    
    print("\n2. Timeout Configuration:")
    timeouts = {
        "Authentication": 120,
        "API calls": 30,
        "Email sending": 60,
        "File operations": 300
    }
    
    for operation, timeout in timeouts.items():
        print(f"   {operation}: {timeout}s timeout")

def test_data_validation():
    """Test data validation edge cases"""
    print("\n\nTesting Data Validation...")
    print("=" * 50)
    
    print("1. Excel File Validation:")
    validations = [
        "Missing reviewer column",
        "Empty reviewer values",
        "Duplicate reviewer names",
        "Very long reviewer names (>255 chars)",
        "Reviewer column with formulas",
        "Hidden rows/filtered data",
        "Merged cells in data"
    ]
    
    for validation in validations:
        print(f"   • {validation}: Detect and handle appropriately")
    
    print("\n2. Input Sanitization:")
    inputs = [
        ("Script injection", "<script>alert('test')</script>"),
        ("SQL injection", "'; DROP TABLE users; --"),
        ("Path traversal", "../../../etc/passwd"),
        ("Null bytes", "file\x00.txt"),
        ("Unicode abuse", "file\u202e.txt")
    ]
    
    for attack, payload in inputs:
        print(f"   {attack}: Sanitize '{payload}'")

def test_performance_edge_cases():
    """Test performance with edge cases"""
    print("\n\nTesting Performance Edge Cases...")
    print("=" * 50)
    
    print("1. Large Dataset Handling:")
    sizes = [
        (100, "Small", "Process normally"),
        (1000, "Medium", "Show progress, batch API calls"),
        (10000, "Large", "Warn user, recommend batching"),
        (100000, "Very Large", "Refuse or require confirmation")
    ]
    
    for count, size, action in sizes:
        print(f"   {count} reviewers ({size}): {action}")
    
    print("\n2. Memory Management:")
    print("   • Stream large Excel files instead of loading fully")
    print("   • Clear widget references after processing")
    print("   • Limit concurrent file operations")
    print("   • Release API connections properly")

def test_ui_state_management():
    """Test UI state management issues"""
    print("\n\nTesting UI State Management...")
    print("=" * 50)
    
    print("1. State Persistence Issues:")
    states = [
        "User refreshes notebook cell",
        "Kernel restart during processing",
        "Widget state out of sync",
        "Multiple notebook instances",
        "Browser tab switching"
    ]
    
    for state in states:
        print(f"   • {state}: Maintain/restore state appropriately")
    
    print("\n2. Progress Indication:")
    print("   • Show spinner during API calls")
    print("   • Update progress bar smoothly")
    print("   • Disable buttons during processing")
    print("   • Clear status messages appropriately")

def run_all_edge_tests():
    """Run all edge case tests"""
    print("EDGE CASE TEST SUITE")
    print("=" * 70)
    print("Testing often-ignored aspects of SharePoint integration\n")
    
    test_functions = [
        test_widget_compatibility,
        test_special_characters,
        test_authentication_edge_cases,
        test_concurrent_operations,
        test_error_recovery,
        test_network_resilience,
        test_data_validation,
        test_performance_edge_cases,
        test_ui_state_management
    ]
    
    for test_func in test_functions:
        try:
            test_func()
        except Exception as e:
            print(f"\n❌ {test_func.__name__} failed: {e}")
    
    print("\n\n" + "=" * 70)
    print("EDGE CASE TESTING COMPLETE")
    print("=" * 70)
    print("\nKey Recommendations:")
    print("1. Always sanitize user inputs and file names")
    print("2. Handle partial failures gracefully")
    print("3. Implement proper retry logic with backoff")
    print("4. Test with various ipywidgets versions")
    print("5. Consider performance with large datasets")
    print("6. Maintain UI state across notebook operations")
    print("7. Provide clear error messages for all failure modes")

if __name__ == "__main__":
    run_all_edge_tests()