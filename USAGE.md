# 📊 Excel Grid Library - Quick Usage Guide

**Created by [SaadCoder](https://saadcoder.com) | [Website](https://saadcoder.com)**

## 🚀 Quick Start

### 1. Include Dependencies
```html
<!-- jQuery (required) -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

<!-- Excel Grid Library -->
<script src="excel-grid.js"></script>
<link rel="stylesheet" href="excel-grid.css">
```

### 2. Basic Usage
```html
<div id="myGrid"></div>

<script>
$('#myGrid').excelGrid({
    rows: 10,
    cols: 8
});
</script>
```

## ⚙️ Configuration Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `rows` | Number | 10 | Number of rows |
| `cols` | Number | 10 | Number of columns |
| `readOnly` | Boolean | false | Make grid read-only |
| `initialData` | Array | null | 2D array to populate |
| `onChange` | Function | null | Change callback |

## 🧮 Formula Examples

### Basic Math
```
=2+2          // 4
=10-5         // 5
=3*4          // 12
=15/3         // 5
```

### Cell References
```
=A1+B1        // Sum of A1 and B1
=A1*B1        // Product of A1 and B1
=(A1+B1)/2    // Average of A1 and B1
```

### Range Functions
```
=SUM(A1:C1)   // Sum of range A1 to C1
=SUM(D2:D5)   // Sum of range D2 to D5
```

## 🎯 API Methods

### Get Data
```javascript
const data = $('#myGrid').data('excelGrid').getData();
console.log(data.raw);        // Raw cell values
console.log(data.evaluated);  // Calculated values
```

### Set Data
```javascript
const newData = [
    ['Name', 'Age', 'Salary'],
    ['John', 30, 50000],
    ['Jane', 25, 45000]
];
$('#myGrid').data('excelGrid').setData(newData);
```

### Destroy Grid
```javascript
$('#myGrid').data('excelGrid').destroy();
```

## ⌨️ Keyboard Shortcuts

- **Click**: Start editing with cursor positioning
- **Double-click**: Start editing with text selection
- **Enter**: Finish editing, stay in same cell
- **Tab**: Finish editing, move right
- **Shift+Tab**: Finish editing, move left
- **Arrow Keys**: Move cursor within text or between cells
- **Escape**: Cancel editing

## 🎨 Custom Styling

```css
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
```

## 🔧 Advanced Usage

### Change Callback
```javascript
$('#myGrid').excelGrid({
    onChange: function(cellCoord, rawValue, evaluatedValue) {
        console.log(`Cell ${cellCoord.ref}: "${rawValue}" → ${evaluatedValue}`);
    }
});
```

### Load Initial Data
```javascript
$('#myGrid').excelGrid({
    initialData: [
        ['Product', 'Price', 'Quantity', 'Total'],
        ['Apple', 1.50, 10, '=B2*C2'],
        ['Banana', 0.80, 15, '=B3*C3']
    ]
});
```

## 🌐 Browser Support

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## 📞 Support

- **Demo**: Open `demo.html` in your browser
- **Documentation**: See `README.md` for full documentation
- **Author**: [SaadCoder](https://saadcoder.com)
- **Website**: [saadcoder.com](https://saadcoder.com)

---

**Made with ❤️ by [SaadCoder](https://saadcoder.com)**
