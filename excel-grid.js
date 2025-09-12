/**
 * Excel Grid Library - A lightweight JavaScript library for creating editable spreadsheet-like grids
 * Dependencies: jQuery
 */

(function($) {
    'use strict';

    // Cell reference utilities
    const CellRef = {
        // Convert A1 notation to row/col indices (0-based)
        parse: function(ref) {
            const match = ref.match(/^([A-Z]+)(\d+)$/);
            if (!match) return null;
            
            const colStr = match[1];
            const row = parseInt(match[2]) - 1;
            
            let col = 0;
            for (let i = 0; i < colStr.length; i++) {
                col = col * 26 + (colStr.charCodeAt(i) - 64);
            }
            col -= 1;
            
            return { row, col };
        },
        
        // Convert row/col indices to A1 notation
        format: function(row, col) {
            let colStr = '';
            col += 1;
            while (col > 0) {
                col -= 1;
                colStr = String.fromCharCode(65 + (col % 26)) + colStr;
                col = Math.floor(col / 26);
            }
            return colStr + (row + 1);
        }
    };

    // Formula parser and evaluator
    const FormulaParser = {
        // Tokenize formula string
        tokenize: function(formula) {
            const tokens = [];
            let i = 0;
            
            while (i < formula.length) {
                const char = formula[i];
                
                if (char === ' ') {
                    i++;
                    continue;
                }
                
                if (char.match(/[+\-*/]/)) {
                    tokens.push({ type: 'operator', value: char });
                    i++;
                } else if (char === '(') {
                    tokens.push({ type: 'paren', value: '(' });
                    i++;
                } else if (char === ')') {
                    tokens.push({ type: 'paren', value: ')' });
                    i++;
                } else if (char.match(/[A-Z]/)) {
                    // Cell reference
                    let ref = '';
                    while (i < formula.length && formula[i].match(/[A-Z0-9]/)) {
                        ref += formula[i];
                        i++;
                    }
                    tokens.push({ type: 'cell', value: ref });
                } else if (char.match(/\d/)) {
                    // Number
                    let num = '';
                    while (i < formula.length && (formula[i].match(/\d/) || formula[i] === '.')) {
                        num += formula[i];
                        i++;
                    }
                    tokens.push({ type: 'number', value: parseFloat(num) });
                } else {
                    throw new Error(`Unexpected character: ${char}`);
                }
            }
            
            return tokens;
        },
        
        // Convert infix to postfix (RPN) using Shunting Yard algorithm
        infixToPostfix: function(tokens) {
            const output = [];
            const operators = [];
            const precedence = { '+': 1, '-': 1, '*': 2, '/': 2 };
            
            for (const token of tokens) {
                if (token.type === 'number' || token.type === 'cell') {
                    output.push(token);
                } else if (token.type === 'operator') {
                    while (operators.length > 0 && 
                           operators[operators.length - 1] !== '(' &&
                           precedence[operators[operators.length - 1]] >= precedence[token.value]) {
                        output.push({ type: 'operator', value: operators.pop() });
                    }
                    operators.push(token.value);
                } else if (token.type === 'paren') {
                    if (token.value === '(') {
                        operators.push('(');
                    } else if (token.value === ')') {
                        while (operators.length > 0 && operators[operators.length - 1] !== '(') {
                            output.push({ type: 'operator', value: operators.pop() });
                        }
                        if (operators.length === 0) {
                            throw new Error('Mismatched parentheses');
                        }
                        operators.pop(); // Remove '('
                    }
                }
            }
            
            while (operators.length > 0) {
                const op = operators.pop();
                if (op === '(') {
                    throw new Error('Mismatched parentheses');
                }
                output.push({ type: 'operator', value: op });
            }
            
            return output;
        },
        
        // Evaluate postfix expression
        evaluate: function(postfix, getCellValue) {
            const stack = [];
            
            for (const token of postfix) {
                if (token.type === 'number') {
                    stack.push(token.value);
                } else if (token.type === 'cell') {
                    const cellValue = getCellValue(token.value);
                    if (cellValue === null || cellValue === undefined) {
                        throw new Error(`Cell ${token.value} not found`);
                    }
                    if (typeof cellValue === 'string' && cellValue.startsWith('#')) {
                        throw new Error(`Cell ${token.value} has error: ${cellValue}`);
                    }
                    stack.push(cellValue);
                } else if (token.type === 'operator') {
                    if (stack.length < 2) {
                        throw new Error('Invalid expression');
                    }
                    
                    const b = stack.pop();
                    const a = stack.pop();
                    
                    // Check for NaN or invalid values
                    if (isNaN(a) || isNaN(b)) {
                        throw new Error('Invalid numeric value in calculation');
                    }
                    
                    switch (token.value) {
                        case '+': stack.push(a + b); break;
                        case '-': stack.push(a - b); break;
                        case '*': stack.push(a * b); break;
                        case '/': 
                            if (b === 0) throw new Error('Division by zero');
                            stack.push(a / b); 
                            break;
                    }
                }
            }
            
            if (stack.length !== 1) {
                throw new Error('Invalid expression');
            }
            
            return stack[0];
        },
        
        // Parse and evaluate formula
        parseAndEvaluate: function(formula, getCellValue) {
            if (!formula.startsWith('=')) {
                throw new Error('Formula must start with =');
            }
            
            const formulaStr = formula.substring(1);
            
            // Handle SUM function
            const sumMatch = formulaStr.match(/^SUM\(([A-Z]+\d+):([A-Z]+\d+)\)$/);
            if (sumMatch) {
                const startRef = sumMatch[1];
                const endRef = sumMatch[2];
                const start = CellRef.parse(startRef);
                const end = CellRef.parse(endRef);
                
                if (!start || !end) {
                    throw new Error('Invalid range in SUM function');
                }
                
                let sum = 0;
                for (let row = start.row; row <= end.row; row++) {
                    for (let col = start.col; col <= end.col; col++) {
                        const value = getCellValue(CellRef.format(row, col));
                        if (typeof value === 'number' && !isNaN(value)) {
                            sum += value;
                        }
                    }
                }
                return sum;
            }
            
            const tokens = this.tokenize(formulaStr);
            const postfix = this.infixToPostfix(tokens);
            return this.evaluate(postfix, getCellValue);
        }
    };

    // Dependency tracker for circular reference detection
    const DependencyTracker = {
        // Build dependency graph
        buildGraph: function(data) {
            const graph = {};
            const dependencies = {};
            
            for (let row = 0; row < data.length; row++) {
                for (let col = 0; col < data[row].length; col++) {
                    const cellRef = CellRef.format(row, col);
                    graph[cellRef] = [];
                    dependencies[cellRef] = [];
                    
                    const value = data[row][col];
                    if (typeof value === 'string' && value.startsWith('=')) {
                        try {
                            const tokens = FormulaParser.tokenize(value.substring(1));
                            for (const token of tokens) {
                                if (token.type === 'cell') {
                                    graph[cellRef].push(token.value);
                                    dependencies[token.value] = dependencies[token.value] || [];
                                    dependencies[token.value].push(cellRef);
                                }
                            }
                        } catch (e) {
                            // Invalid formula, skip dependencies
                        }
                    }
                }
            }
            
            return { graph, dependencies };
        },
        
        // Detect circular references using DFS
        detectCircular: function(graph, startCell) {
            const visited = new Set();
            const recursionStack = new Set();
            
            const dfs = (cell) => {
                if (recursionStack.has(cell)) {
                    return true; // Circular reference found
                }
                
                if (visited.has(cell)) {
                    return false;
                }
                
                visited.add(cell);
                recursionStack.add(cell);
                
                for (const neighbor of graph[cell] || []) {
                    if (dfs(neighbor)) {
                        return true;
                    }
                }
                
                recursionStack.delete(cell);
                return false;
            };
            
            return dfs(startCell);
        },
        
        // Get cells that depend on a given cell (for recalculation)
        getDependents: function(dependencies, cell) {
            if (!dependencies) return [];
            
            const dependents = new Set();
            const queue = [cell];
            
            while (queue.length > 0) {
                const current = queue.shift();
                for (const dependent of dependencies[current] || []) {
                    if (!dependents.has(dependent)) {
                        dependents.add(dependent);
                        queue.push(dependent);
                    }
                }
            }
            
            return Array.from(dependents);
        },
        
        // Topological sort for proper evaluation order
        topologicalSort: function(graph) {
            const visited = new Set();
            const temp = new Set();
            const result = [];
            
            const visit = (cell) => {
                if (temp.has(cell)) {
                    throw new Error('Circular dependency detected');
                }
                if (visited.has(cell)) {
                    return;
                }
                
                temp.add(cell);
                for (const neighbor of graph[cell] || []) {
                    visit(neighbor);
                }
                temp.delete(cell);
                visited.add(cell);
                result.push(cell);
            };
            
            for (const cell in graph) {
                if (!visited.has(cell)) {
                    visit(cell);
                }
            }
            
            return result.reverse(); // Reverse to get dependency order
        }
    };

    // Main ExcelGrid class
    function ExcelGrid(containerOrTable, options) {
        this.options = $.extend({
            rows: 10,
            cols: 10,
            readOnly: false,
            initialData: null,
            onChange: null
        }, options || {});
        
        this.container = $(containerOrTable);
        this.data = [];
        this.evaluatedData = [];
        this.dependencies = null;
        this.graph = null;
        this.editingCell = null;
        this.originalTable = null;
        
        this.init();
    }
    
    ExcelGrid.prototype = {
        init: function() {
            this.setupData();
            this.updateDependencies();
            this.createGrid();
            this.bindEvents();
        },
        
        setupData: function() {
            if (this.options.initialData) {
                this.data = this.options.initialData.map(row => [...row]);
            } else {
                this.data = Array(this.options.rows).fill().map(() => Array(this.options.cols).fill(''));
            }
            
            this.evaluatedData = this.data.map(row => [...row]);
            this.evaluateAll();
        },
        
        createGrid: function() {
            if (this.container.is('table')) {
                // Replace existing table
                this.originalTable = this.container.clone();
                this.container.empty();
            } else {
                // Create new table in container
                this.container.empty();
            }
            
            const table = $('<table class="excel-grid"></table>');
            const tbody = $('<tbody></tbody>');
            
            for (let row = 0; row < this.data.length; row++) {
                const tr = $('<tr></tr>');
                
                for (let col = 0; col < this.data[row].length; col++) {
                    const cellRef = CellRef.format(row, col);
                    const td = $(`<td data-row="${row}" data-col="${col}" data-ref="${cellRef}"></td>`);
                    
                    this.updateCellDisplay(td, row, col);
                    tr.append(td);
                }
                
                tbody.append(tr);
            }
            
            table.append(tbody);
            this.container.append(table);
        },
        
        updateCellDisplay: function(td, row, col) {
            const value = this.data[row][col];
            const evaluatedValue = this.evaluatedData[row][col];
            
            if (typeof value === 'string' && value.startsWith('=')) {
                // Formula cell
                if (this.isError(evaluatedValue)) {
                    td.addClass('error').text(evaluatedValue);
                } else {
                    td.removeClass('error').text(evaluatedValue);
                }
            } else {
                // Regular cell
                td.removeClass('error').text(value);
            }
        },
        
        isError: function(value) {
            return typeof value === 'string' && value.startsWith('#');
        },
        
        bindEvents: function() {
            const self = this;
            
            // Cell click events - single click for cursor positioning
            this.container.on('click', 'td', function(e) {
                if (self.options.readOnly) return;
                
                const row = parseInt($(this).data('row'));
                const col = parseInt($(this).data('col'));
                self.startEdit(row, col, false); // false = don't select all
            });
            
            // Double click to edit with text selection
            this.container.on('dblclick', 'td', function(e) {
                if (self.options.readOnly) return;
                
                const row = parseInt($(this).data('row'));
                const col = parseInt($(this).data('col'));
                self.startEdit(row, col, true); // true = select all text
            });
            
            // Keyboard navigation - FIXED: Arrow keys should move cursor within text, not between cells
            this.container.on('keydown', 'input', function(e) {
                const row = parseInt($(this).data('row'));
                const col = parseInt($(this).data('col'));
                
                switch (e.key) {
                    case 'Enter':
                        e.preventDefault();
                        self.finishEdit();
                        // FIXED: Enter should stay in same cell, not move to next cell
                        break;
                    case 'Escape':
                        e.preventDefault();
                        self.cancelEdit();
                        break;
                    case 'Tab':
                        e.preventDefault();
                        self.finishEdit();
                        if (e.shiftKey) {
                            // Move to previous column
                            if (col > 0) {
                                self.startEdit(row, col - 1, true);
                            }
                        } else {
                            // Move to next column
                            if (col + 1 < self.data[0].length) {
                                self.startEdit(row, col + 1, true);
                            }
                        }
                        break;
                    case 'ArrowUp':
                        // FIXED: Don't prevent default - let arrow keys move cursor within text
                        // Only move between cells if cursor is at the beginning/end and we're not editing
                        if (this.selectionStart === 0 && this.selectionEnd === 0) {
                            e.preventDefault();
                            self.finishEdit();
                            if (row > 0) {
                                self.startEdit(row - 1, col, true);
                            }
                        }
                        break;
                    case 'ArrowDown':
                        // FIXED: Don't prevent default - let arrow keys move cursor within text
                        // Only move between cells if cursor is at the end and we're not editing
                        if (this.selectionStart === this.value.length && this.selectionEnd === this.value.length) {
                            e.preventDefault();
                            self.finishEdit();
                            if (row + 1 < self.data.length) {
                                self.startEdit(row + 1, col, true);
                            }
                        }
                        break;
                    case 'ArrowLeft':
                        // FIXED: Don't prevent default - let arrow keys move cursor within text
                        // Only move between cells if cursor is at the beginning and we're not editing
                        if (this.selectionStart === 0 && this.selectionEnd === 0) {
                            e.preventDefault();
                            self.finishEdit();
                            if (col > 0) {
                                self.startEdit(row, col - 1, true);
                            }
                        }
                        break;
                    case 'ArrowRight':
                        // FIXED: Don't prevent default - let arrow keys move cursor within text
                        // Only move between cells if cursor is at the end and we're not editing
                        if (this.selectionStart === this.value.length && this.selectionEnd === this.value.length) {
                            e.preventDefault();
                            self.finishEdit();
                            if (col + 1 < self.data[0].length) {
                                self.startEdit(row, col + 1, true);
                            }
                        }
                        break;
                }
            });
            
            // Blur event to finish editing
            this.container.on('blur', 'input', function() {
                self.finishEdit();
            });
        },
        
        startEdit: function(row, col, selectAll = false) {
            if (this.editingCell) {
                this.finishEdit();
            }
            
            const cell = this.container.find(`td[data-row="${row}"][data-col="${col}"]`);
            const currentValue = this.data[row][col];
            
            const input = $('<input type="text">')
                .val(currentValue)
                .data('row', row)
                .data('col', col)
                .addClass('cell-input');
            
            cell.html(input);
            input.focus();
            
            // Select all text only if requested (double-click) or if cell is empty
            if (selectAll || !currentValue) {
                input.select();
            } else {
                // Position cursor at the end for single-click editing
                const length = currentValue.length;
                input[0].setSelectionRange(length, length);
            }
            
            this.editingCell = { row, col, input };
        },
        
        finishEdit: function() {
            if (!this.editingCell) return;
            
            const { row, col, input } = this.editingCell;
            const newValue = input.val();
            
            this.setCellValue(row, col, newValue);
            this.editingCell = null;
        },
        
        cancelEdit: function() {
            if (!this.editingCell) return;
            
            const { row, col } = this.editingCell;
            const cell = this.container.find(`td[data-row="${row}"][data-col="${col}"]`);
            
            this.updateCellDisplay(cell, row, col);
            this.editingCell = null;
        },
        
        moveToCell: function(row, col) {
            if (row < 0 || row >= this.data.length || col < 0 || col >= this.data[0].length) {
                return;
            }
            
            // Start editing the target cell directly
            this.startEdit(row, col, true);
        },
        
        setCellValue: function(row, col, value) {
            const oldValue = this.data[row][col];
            this.data[row][col] = value;
            
            // Update dependencies
            this.updateDependencies();
            
            // Recalculate affected cells
            this.recalculateCell(row, col);
            
            // Update display
            const cell = this.container.find(`td[data-row="${row}"][data-col="${col}"]`);
            this.updateCellDisplay(cell, row, col);
            
            // Trigger change callback
            if (this.options.onChange) {
                this.options.onChange(
                    { row, col, ref: CellRef.format(row, col) },
                    value,
                    this.evaluatedData[row][col]
                );
            }
        },
        
        updateDependencies: function() {
            const result = DependencyTracker.buildGraph(this.data);
            this.graph = result.graph;
            this.dependencies = result.dependencies;
        },
        
        evaluateCell: function(row, col) {
            const cellRef = CellRef.format(row, col);
            const value = this.data[row][col];
            
            try {
                if (typeof value === 'string' && value.startsWith('=')) {
                    // Check for circular references
                    if (DependencyTracker.detectCircular(this.graph, cellRef)) {
                        this.evaluatedData[row][col] = '#CIRC';
                        return;
                    }
                    
                    // Evaluate formula
                    const getCellValue = (ref) => {
                        const parsed = CellRef.parse(ref);
                        if (!parsed) return null;
                        
                        // Check if the referenced cell is within bounds
                        if (parsed.row < 0 || parsed.row >= this.data.length || 
                            parsed.col < 0 || parsed.col >= this.data[0].length) {
                            return null;
                        }
                        
                        const cellValue = this.evaluatedData[parsed.row] && this.evaluatedData[parsed.row][parsed.col];
                        
                        // If the cell hasn't been evaluated yet, try to evaluate it now
                        if (cellValue === undefined && this.data[parsed.row] && this.data[parsed.row][parsed.col]) {
                            const cellData = this.data[parsed.row][parsed.col];
                            if (typeof cellData === 'string' && cellData.startsWith('=')) {
                                // This is a formula cell that hasn't been evaluated yet
                                // We need to evaluate it first
                                this.evaluateCell(parsed.row, parsed.col);
                                return this.evaluatedData[parsed.row][parsed.col];
                            }
                        }
                        
                        return cellValue;
                    };
                    
                    this.evaluatedData[row][col] = FormulaParser.parseAndEvaluate(value, getCellValue);
                } else {
                    // Regular value
                    this.evaluatedData[row][col] = value;
                }
            } catch (error) {
                console.log(`Error evaluating cell ${cellRef}: ${error.message}`);
                this.evaluatedData[row][col] = '#ERROR';
            }
        },
        
        recalculateCell: function(row, col) {
            this.evaluateCell(row, col);
            
            // Update display if grid exists
            if (this.container.find('td').length > 0) {
                const cell = this.container.find(`td[data-row="${row}"][data-col="${col}"]`);
                this.updateCellDisplay(cell, row, col);
            }
            
            // Recalculate dependent cells
            const cellRef = CellRef.format(row, col);
            const dependents = this.dependencies ? DependencyTracker.getDependents(this.dependencies, cellRef) : [];
            for (const dependentRef of dependents) {
                const parsed = CellRef.parse(dependentRef);
                if (parsed) {
                    this.recalculateCell(parsed.row, parsed.col);
                }
            }
        },
        
        evaluateAll: function() {
            // First pass: evaluate all cells
            for (let row = 0; row < this.data.length; row++) {
                for (let col = 0; col < this.data[row].length; col++) {
                    this.evaluateCell(row, col);
                }
            }
        },
        
        getData: function() {
            return {
                raw: this.data.map(row => [...row]),
                evaluated: this.evaluatedData.map(row => [...row])
            };
        },
        
        setData: function(data) {
            this.data = data.map(row => [...row]);
            this.evaluatedData = this.data.map(row => [...row]);
            this.updateDependencies();
            this.evaluateAll();
            this.createGrid();
        },
        
        destroy: function() {
            this.container.off();
            if (this.originalTable) {
                this.container.replaceWith(this.originalTable);
            } else {
                this.container.empty();
            }
        }
    };
    
    // jQuery plugin - Fixed to properly store and return the instance
    $.fn.excelGrid = function(options) {
        const instances = [];
        this.each(function() {
            const $this = $(this);
            const instance = new ExcelGrid(this, options);
            $this.data('excelGrid', instance);
            instances.push(instance);
        });
        return instances.length === 1 ? instances[0] : instances;
    };
    
    // Expose ExcelGrid globally
    window.ExcelGrid = ExcelGrid;
    
})(jQuery);
