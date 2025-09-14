# 📊 Excel Grid Library

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![jQuery](https://img.shields.io/badge/jQuery-3.6.0-blue.svg)](https://jquery.com/)
[![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-yellow.svg)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![Browser Support](https://img.shields.io/badge/Browser-Chrome%20%7C%20Firefox%20%7C%20Safari%20%7C%20Edge-green.svg)](https://caniuse.com/)

A lightweight, feature-rich JavaScript library that transforms HTML tables or 2D arrays into fully functional, Excel-like editable grids with formula support, dependency tracking, and intuitive keyboard navigation.
👉 **Live Demo:** [Excel Grid Library Demo](https://saadcoder.com/saadapps/Excel-Grid-Library/demo.html)

**Created by [SaadCoder](https://saadcoder.com) | [![Website](https://img.shields.io/badge/Website-saadcoder.com-brightgreen.svg)](https://saadcoder.com)**

## ✨ Features

### 🎯 Core Functionality
- **Editable Grid**: Convert any HTML `<table>` or 2D array into an interactive spreadsheet
- **Inline Editing**: Click to edit cells with smart cursor positioning
- **Formula Support**: Full Excel-style formulas with `=` prefix
- **Cell References**: A1-style references (A1, B2, AA1, etc.)
- **Real-time Calculation**: Automatic recalculation with dependency tracking
- **Error Handling**: Circular reference detection and formula error reporting

### 🧮 Formula Engine
- **Arithmetic Operations**: `+`, `-`, `*`, `/` with proper operator precedence
- **Parentheses Support**: Complex expressions with `()` grouping
- **Built-in Functions**: `SUM(range)` for range calculations
- **Cell References**: Excel-style references up to ZZ columns
- **Error Detection**: `#ERROR`, `#CIRC` for invalid formulas and circular references

### ⌨️ Excel-like Navigation
- **Smart Arrow Keys**: Move cursor within text when editing, navigate cells when at boundaries
- **Enter Key**: Finish editing and stay in current cell
- **Tab Navigation**: Move between cells horizontally
- **Escape**: Cancel editing without saving changes
- **Click Behavior**: Single-click for cursor positioning, double-click for text selection

### 🎨 User Experience
- **Responsive Design**: Works on desktop and mobile devices
- **Clean Interface**: Minimal, professional styling
- **Keyboard Shortcuts**: Full Excel-like keyboard support
- **Copy/Paste**: Basic text clipboard support
- **Performance**: Optimized for grids up to 200x200 cells

## 🚀 Quick Start

### Installation

```html
<!-- Include jQuery (required) -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

<!-- Include Excel Grid Library -->
<script src="excel-grid.js"></script>
<link rel="stylesheet" href="excel-grid.css">
```

### Basic Usage

```html
<!-- HTML -->
<div id="myGrid"></div>

<script>
// Initialize with default options
$('#myGrid').excelGrid();

// Or with custom options
$('#myGrid').excelGrid({
    rows: 10,
    cols: 8,
    onChange: function(cellCoord, rawValue, evaluatedValue) {
        console.log('Cell changed:', cellCoord, rawValue, evaluatedValue);
    }
});
</script>
```

### Load Data Programmatically

```javascript
const data = [
    ['Product', 'Price', 'Quantity', 'Total'],
    ['Apple', 1.50, 10, '=B2*C2'],
    ['Banana', 0.80, 15, '=B3*C3'],
    ['Total', '', '', '=D2+D3']
];

$('#myGrid').excelGrid({
    initialData: data
});
```

## 📖 API Reference

### Initialization Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `rows` | Number | 10 | Number of rows in the grid |
| `cols` | Number | 10 | Number of columns in the grid |
| `readOnly` | Boolean | false | Make the grid read-only |
| `initialData` | Array | null | 2D array to populate the grid |
| `onChange` | Function | null | Callback when cell values change |

### Methods

#### `getData()`
Returns current grid data as an object with `raw` and `evaluated` properties.

```javascript
const data = $('#myGrid').data('excelGrid').getData();
console.log('Raw data:', data.raw);
console.log('Evaluated data:', data.evaluated);
```

#### `setData(data)`
Load new data into the grid.

```javascript
const newData = [
    ['Name', 'Age', 'Salary'],
    ['John', 30, 50000],
    ['Jane', 25, 45000]
];

$('#myGrid').data('excelGrid').setData(newData);
```

#### `destroy()`
Remove the grid and restore original DOM.

```javascript
$('#myGrid').data('excelGrid').destroy();
```

### Events

#### `onChange(cellCoord, rawValue, evaluatedValue)`
Triggered when a cell value changes.

```javascript
$('#myGrid').excelGrid({
    onChange: function(cellCoord, rawValue, evaluatedValue) {
        console.log(`Cell ${cellCoord.ref}: "${rawValue}" → ${evaluatedValue}`);
    }
});
```

## 🧮 Formula Examples

### Basic Arithmetic
```
=2+2                    // 4
=10-5                   // 5
=3*4                    // 12
=15/3                   // 5
```

### Cell References
```
=A1+B1                  // Sum of A1 and B1
=A1*B1                  // Product of A1 and B1
=(A1+B1)/2              // Average of A1 and B1
```

### Complex Formulas
```
=(A1+B1)*C1             // (A1+B1) multiplied by C1
=A1*(B1-C1)             // A1 multiplied by (B1-C1)
=(A1+B1)/(C1-D1)        // (A1+B1) divided by (C1-D1)
```

### Range Functions
```
=SUM(A1:C1)             // Sum of range A1 to C1
=SUM(D2:D5)             // Sum of range D2 to D5
```

### Error Cases
```
=A1/B1                  // #ERROR if B1 is 0
=A1+B1                  // #ERROR if A1 or B1 is empty
=A1                     // #CIRC if A1 references itself
```

## 🎨 Styling

The library includes minimal CSS that you can customize:

```css
.excel-grid {
    border-collapse: collapse;
    width: 100%;
}

.excel-grid td {
    border: 1px solid #ddd;
    padding: 8px;
    min-width: 80px;
    height: 30px;
}

.excel-grid td:hover {
    background-color: #f5f5f5;
}

.excel-grid td.error {
    color: #d32f2f;
    background-color: #ffebee;
}

.cell-input {
    width: 100%;
    height: 100%;
    border: none;
    outline: none;
    background: transparent;
}
```

## 🔧 Advanced Usage

### Custom Formula Functions

The library supports extending with custom functions by modifying the `FormulaParser.parseAndEvaluate` method.

### Performance Optimization

For large grids (200x200+), consider:
- Debouncing the `onChange` callback
- Implementing virtual scrolling for very large datasets
- Using `requestAnimationFrame` for smooth updates

### Integration with Frameworks

#### React
```jsx
import React, { useEffect, useRef } from 'react';

function ExcelGridComponent() {
    const gridRef = useRef(null);
    
    useEffect(() => {
        if (gridRef.current) {
            $(gridRef.current).excelGrid({
                rows: 10,
                cols: 8
            });
        }
    }, []);
    
    return <div ref={gridRef}></div>;
}
```

#### Vue.js
```vue
<template>
    <div ref="gridContainer"></div>
</template>

<script>
export default {
    mounted() {
        this.$nextTick(() => {
            $(this.$refs.gridContainer).excelGrid({
                rows: 10,
                cols: 8
            });
        });
    }
}
</script>
```

## 🌐 Browser Support

- **Chrome** 60+
- **Firefox** 55+
- **Safari** 12+
- **Edge** 79+

## 📦 File Structure

```
excel-clone/
├── excel-grid.js          # Main library file
├── excel-grid.css         # Default styles
├── demo.html             # Interactive demo
├── package.json          # NPM package info
├── USAGE.md              # Quick usage guide
└── README.md             # This file
```

## 🤝 Contributing

We welcome contributions! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### Development Setup

1. Clone the repository
2. Open `demo.html` in your browser
3. Make changes to `excel-grid.js`
4. Test your changes in the demo
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [jQuery](https://jquery.com/)
- Inspired by Microsoft Excel's functionality
- Formula parsing uses the Shunting Yard algorithm
- Dependency tracking uses topological sorting
- Created by [SaadCoder](https://saadcoder.com)

## 📞 Support

If you encounter any issues or have questions:

1. Check the [demo page](demo.html) for examples
2. Review the [USAGE.md](USAGE.md) for quick reference
3. Open an issue on GitHub
4. Check existing issues for solutions
5. Visit [saadcoder.com](https://saadcoder.com) for more projects

---

**Made with ❤️ by [SaadCoder](https://saadcoder.com)**

*Transform your data tables into powerful, interactive spreadsheets with just a few lines of code!*
