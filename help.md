## String Manipulation Functions in Python (VB6 to Python Translation)

### Purpose

This set of functions provides a collection of string manipulation utilities that are commonly used in VB6 programming. These functions are translated into their equivalent Python counterparts, allowing VB6 developers to seamlessly transition to Python programming.

### Usage

The functions can be imported into any Python module using the following statement:

from string_manipulation import *

Once imported, the functions can be called directly within your Python code.

### Function Descriptions

#### `Left(s, n)`

* **Purpose:** Returns the first `n` characters of the string `s`.

* **VB6 Equivalent:** `Left(s, n)`

* **Python Usage:**

left_string = Left("Hello, World!", 5)
print(left_string)  # Output: "Hello"

#### `Right(s, n)`

* **Purpose:** Returns the last `n` characters of the string `s`.

* **VB6 Equivalent:** `Right(s, n)`

* **Python Usage:**

right_string = Right("Hello, World!", 5)
print(right_string)  # Output: "World!"

#### `Mid(s, start, length=None)`

* **Purpose:** Returns a substring of `length` characters starting from position `start` in the string `s`.

* **VB6 Equivalent:** `Mid(s, start, length)`

* **Python Usage:**

mid_string = Mid("Hello, World!", 7, 5)
print(mid_string)  # Output: "World"

#### `Len(s)`

* **Purpose:** Returns the length of the string `s`.

* **VB6 Equivalent:** `Len(s)`

* **Python Usage:**

string_length = Len("Hello, World!")
print(string_length)  # Output: 13

#### `InStr(s, substring, start=0)`

* **Purpose:** Returns the position of the first occurrence of the substring `substring` in the string `s`, starting the search from position `start`. Returns 0 if not found.

* **VB6 Equivalent:** `InStr(s, substring, start)`

* **Python Usage:**

substring_position = InStr("Hello, World!", "World", 6)
print(substring_position)  # Output: 7

#### `Trim(s)`

* **Purpose:** Removes leading and trailing whitespace from the string `s`.

* **VB6 Equivalent:** `Trim(s)`

* **Python Usage:**

trimmed_string = Trim(" Hello, World! ")
print(trimmed_string)  # Output: "Hello, World!"

#### `LTrim(s)`

* **Purpose:** Removes leading whitespace from the string `s`.

* **VB6 Equivalent:** `LTrim(s)`

* **Python Usage:**

ltrimmed_string = LTrim("   Hello, World!")
print(ltrimmed_string)  # Output: "Hello, World!"

#### `RTrim(s)`

* **Purpose:** Removes trailing whitespace from the string `s`.

* **VB6 Equivalent:** `RTrim(s)`

* **Python Usage:**

rtrimmed_string = RTrim("Hello, World!    ")
print(rtrimmed_string)  # Output: "Hello, World!"

#### `LCase(s)`

* **Purpose:** Converts the string `s` to lowercase.

* **VB6 Equivalent:** `LCase(s)`

* **Python Usage:**

lowercase_string = LCase("HELLO, WORLD!")
print(lowercase_string)  # Output: "hello, world!"

#### `UCase(s)`

* **Purpose:** Converts the string `s` to uppercase.

* **VB6 Equivalent:** `UCase(s)`

* **Python Usage:**

uppercase_string = UCase("hello, world!")
print(uppercase_string)  # Output: "HELLO, WORLD!"

#### `Capitalize(s, sep=' ')`

* **Purpose:** Capitalizes the first letter of each word in the string `s`, separated by the delimiter `sep`.

* **VB6 Equivalent:** `Capitalize(s, sep)`

* **Python Usage:**

capitalized_string = Capitalize
