import json
import sys
from typing import Any, Dict, Union, List


def _convert_to_list_of_dicts(data: Any) -> List[Dict[str, Any]]:
    """
    Convert various data formats to a list of dictionaries for Excel compatibility.
    
    Args:
        data: The data to convert
        
    Returns:
        List[Dict[str, Any]]: Data formatted as list of dictionaries
        
    Raises:
        ValueError: If data cannot be converted to list of dictionaries
    """
    if data is None:
        return []
    
    # Already a list of dictionaries
    if isinstance(data, list):
        if all(isinstance(item, dict) for item in data):
            return data
        elif all(isinstance(item, (str, int, float, bool)) for item in data):
            # Convert list of primitives to list of dicts with 'value' key
            return [{"value": item} for item in data]
        else:
            # Convert mixed list to list of dicts
            result = []
            for i, item in enumerate(data):
                if isinstance(item, dict):
                    result.append(item)
                else:
                    result.append({"value": item, "index": i})
            return result
    
    # Single dictionary - wrap in list
    elif isinstance(data, dict):
        # If it looks like it has multiple records (keys are indices or similar)
        if all(isinstance(key, (int, str)) and isinstance(value, dict) for key, value in data.items()):
            # Convert {0: {data}, 1: {data}} format to list of dicts
            try:
                sorted_items = sorted(data.items(), key=lambda x: int(x[0]))
                return [value for key, value in sorted_items]
            except (ValueError, TypeError):
                # If keys aren't numeric, add key as a field
                return [{**value, "key": key} for key, value in data.items()]
        else:
            # Single record dictionary
            return [data]
    
    # Primitive value
    elif isinstance(data, (str, int, float, bool)):
        return [{"value": data}]
    
    else:
        raise ValueError(f"Cannot convert data of type {type(data)} to list of dictionaries")


def execute_python_with_json(python_code: str, json_data: Union[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Execute Python code with a JSON object available as context.
    Always returns a list of dictionaries suitable for Excel export.
    
    Args:
        python_code (str): The Python code to execute as a string
        json_data (Union[str, Dict[str, Any]]): JSON data as a string or dictionary
        
    Returns:
        List[Dict[str, Any]]: The result as a list of dictionaries
        
    Raises:
        Exception: If there's an error in code execution or JSON parsing
        ValueError: If the result cannot be converted to list of dictionaries
    """
    try:
        # Parse JSON if it's a string
        if isinstance(json_data, str):
            data = json.loads(json_data)
        else:
            data = json_data
            
        # Create a safe execution environment
        # Include common modules that might be useful
        execution_globals = {
            '__builtins__': __builtins__,
            'json': json,
            'data': data,  # Make JSON data available as 'data' variable
            'result': None  # Variable to store the result
        }
        
        # Execute the code
        exec(python_code, execution_globals)
        
        # Get the result
        raw_result = execution_globals.get('result')
        
        # Convert to list of dictionaries
        return _convert_to_list_of_dicts(raw_result)
        
    except json.JSONDecodeError as e:
        raise Exception(f"Invalid JSON data: {e}")
    except SyntaxError as e:
        raise Exception(f"Invalid Python code syntax: {e}")
    except ValueError as e:
        raise Exception(f"Result validation error: {e}")
    except Exception as e:
        raise Exception(f"Error executing Python code: {e}")


def evaluate_python_expression(expression: str, json_data: Union[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Evaluate a Python expression with JSON data as context.
    Always returns a list of dictionaries suitable for Excel export.
    
    Args:
        expression (str): The Python expression to evaluate
        json_data (Union[str, Dict[str, Any]]): JSON data as a string or dictionary
        
    Returns:
        List[Dict[str, Any]]: The result as a list of dictionaries
    """
    try:
        # Parse JSON if it's a string
        if isinstance(json_data, str):
            data = json.loads(json_data)
        else:
            data = json_data
            
        # Create evaluation context
        eval_globals = {
            '__builtins__': {'len': len, 'sum': sum, 'max': max, 'min': min, 'abs': abs, 'round': round},
            'json': json,
            'data': data
        }
        
        # Evaluate the expression
        raw_result = eval(expression, eval_globals)
        
        # Convert to list of dictionaries
        return _convert_to_list_of_dicts(raw_result)
        
    except json.JSONDecodeError as e:
        raise Exception(f"Invalid JSON data: {e}")
    except SyntaxError as e:
        raise Exception(f"Invalid Python expression syntax: {e}")
    except ValueError as e:
        raise Exception(f"Result validation error: {e}")
    except Exception as e:
        raise Exception(f"Error evaluating Python expression: {e}")


def validate_excel_format(data: List[Dict[str, Any]]) -> bool:
    """
    Validate that data is in the correct format for Excel export.
    
    Args:
        data: The data to validate
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not isinstance(data, list):
        return False
    
    if not data:  # Empty list is valid
        return True
        
    # Check if all items are dictionaries
    if not all(isinstance(item, dict) for item in data):
        return False
        
    # Check if all dictionaries have string keys (required for Excel headers)
    for item in data:
        if not all(isinstance(key, str) for key in item.keys()):
            return False
            
    return True


# Example usage and testing
if __name__ == "__main__":
    # Example 1: Code that returns a list of dictionaries
    sample_json = {
        "students": [
            {"name": "Alice", "math": 95, "science": 87},
            {"name": "Bob", "math": 78, "science": 92},
            {"name": "Charlie", "math": 88, "science": 85}
        ],
        "multiplier": 1.1
    }
    
    sample_code_1 = """
# Process student data and add calculated fields
result = []
for student in data['students']:
    processed_student = {
        'name': student['name'],
        'math_score': student['math'],
        'science_score': student['science'],
        'average_score': (student['math'] + student['science']) / 2,
        'boosted_average': ((student['math'] + student['science']) / 2) * data['multiplier'],
        'grade': 'A' if (student['math'] + student['science']) / 2 >= 90 else 'B' if (student['math'] + student['science']) / 2 >= 80 else 'C'
    }
    result.append(processed_student)
"""
    
    try:
        result = execute_python_with_json(sample_code_1, sample_json)
        print("Example 1 - Student processing result:")
        print(f"Is valid Excel format: {validate_excel_format(result)}")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(f"Error: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Example 2: Code that returns a single dictionary (will be converted to list)
    sample_code_2 = """
result = {
    'total_students': len(data['students']),
    'avg_math': sum(s['math'] for s in data['students']) / len(data['students']),
    'avg_science': sum(s['science'] for s in data['students']) / len(data['students'])
}
"""
    
    try:
        result = execute_python_with_json(sample_code_2, sample_json)
        print("Example 2 - Summary statistics (single dict converted to list):")
        print(f"Is valid Excel format: {validate_excel_format(result)}")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(f"Error: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Example 3: Expression that returns a list
    expression = "[{'student': s['name'], 'total': s['math'] + s['science']} for s in data['students']]"
    
    try:
        result = evaluate_python_expression(expression, sample_json)
        print("Example 3 - Expression result:")
        print(f"Is valid Excel format: {validate_excel_format(result)}")
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(f"Error: {e}")
