import java.awt.Color;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.regex.Matcher;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;
import java.util.regex.Pattern;
import javax.swing.table.AbstractTableModel;

/**
 * Data-model layer: one worksheet grid (default 60 rows x 26 columns, headers A-Z).
 *
 * Three separate maps keyed by cell reference ("A1", "B7", ...):
 * <ul>
 *   <li>{@code rawValues} - literal text / numbers typed by the user</li>
 *   <li>{@code formulas}  - inputs that start with '='</li>
 *   <li>{@code styles}    - bold / italic / alignment / colours / font size</li>
 * </ul>
 * A fourth map caches computed formula results. Any edit clears the cache and
 * fires a table update, so every dependent formula is recalculated lazily the
 * next time it is painted or referenced (circular references yield #ERROR!).
 */
public class SpreadsheetModel extends AbstractTableModel {

    public static final int DEFAULT_ROWS = 60;
    public static final int DEFAULT_COLS = 26;

    /** Column widths are stored at 100% zoom, in pixels. */
    public static final int DEFAULT_COLUMN_WIDTH = 90;
    public static final int MIN_COLUMN_WIDTH = 20;
    public static final int MAX_COLUMN_WIDTH = 800;
    public static final int MIN_ZOOM = 25;
    public static final int MAX_ZOOM = 300;

    private static final Pattern NUMBER =
            Pattern.compile("^[-+]?(\\d+(\\.\\d*)?|\\.\\d+)([eE][-+]?\\d+)?$");

    private static final Pattern PERCENT_INPUT = Pattern.compile("^([-+]?(\\d+(\\.\\d*)?|\\.\\d+))%$");

    private final int rows;
    private final int cols;

    private final Map<String, String> rawValues = new HashMap<>();
    private final Map<String, String> formulas = new HashMap<>();
    private final Map<String, CellStyle> styles = new HashMap<>();
    private final Map<String, Object> computed = new HashMap<>();
    private final Set<String> evaluating = new HashSet<>();

    private final FormulaEngine engine;

    /** Sheet protection state and the pictures floating above the grid. */
    private final SheetProtection protection = new SheetProtection();
    private final List<SheetImage> images = new ArrayList<>();
    private final List<DrawShape> shapes = new ArrayList<>();
    private final List<ChartSpec> charts = new ArrayList<>();

    /** View settings kept with the sheet (and saved in the workbook file). */
    private final int[] columnWidths;
    private int zoomPercent = 100;

    /** Conditional formatting, the sheet filter, and a view flag (not saved) that shows formulas instead of values. */
    private final List<CondRule> condRules = new ArrayList<>();
    private final Map<CondRule, Map<String, Integer>> dupCache = new HashMap<>();
    private AutoFilter filter;
    private boolean showFormulas;

    private int batchDepth = 0;
    private boolean dirty = false;

    public SpreadsheetModel() {
        this(DEFAULT_ROWS, DEFAULT_COLS);
    }

    public SpreadsheetModel(int rows, int cols) {
        this.rows = rows;
        this.cols = cols;
        this.columnWidths = new int[cols];
        java.util.Arrays.fill(columnWidths, DEFAULT_COLUMN_WIDTH);
        this.engine = new FormulaEngine(this::getComputedValue, rows, cols);
    }

    // ------------------------------------------------------------------
    // TableModel contract
    // ------------------------------------------------------------------

    @Override
    public int getRowCount() {
        return rows;
    }

    @Override
    public int getColumnCount() {
        return cols;
    }

    @Override
    public String getColumnName(int column) {
        return FormulaEngine.columnLabel(column);
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        return Object.class;
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return !protection.isProtected();
    }

    // ------------------------------------------------------------------
    // View settings: column widths and zoom (not affected by protection)
    // ------------------------------------------------------------------

    /** Width of a column in pixels at 100% zoom. */
    public int getColumnWidth(int col) {
        return col >= 0 && col < cols ? columnWidths[col] : DEFAULT_COLUMN_WIDTH;
    }

    public void setColumnWidth(int col, int width) {
        if (col >= 0 && col < cols) {
            columnWidths[col] = Math.max(MIN_COLUMN_WIDTH, Math.min(MAX_COLUMN_WIDTH, width));
        }
    }

    public void resetColumnWidths() {
        java.util.Arrays.fill(columnWidths, DEFAULT_COLUMN_WIDTH);
    }

    public int getZoomPercent() {
        return zoomPercent;
    }

    public void setZoomPercent(int percent) {
        zoomPercent = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, percent));
    }

    // ------------------------------------------------------------------
    // Protection and pictures
    // ------------------------------------------------------------------

    public SheetProtection getProtection() {
        return protection;
    }

    public boolean isProtected() {
        return protection.isProtected();
    }

    /** Read-only view of the freehand strokes and shapes, back to front. */
    public List<DrawShape> getShapes() {
        return Collections.unmodifiableList(shapes);
    }

    public void addShape(DrawShape shape) {
        if (!protection.isProtected()) shapes.add(shape);
    }

    public void removeShape(DrawShape shape) {
        if (!protection.isProtected()) shapes.remove(shape);
    }

    /** Replaces all drawings (used by undo / redo and clear). */
    public void setShapes(List<DrawShape> list) {
        if (protection.isProtected()) return;
        shapes.clear();
        shapes.addAll(list);
    }

    /** Read-only view of the pictures, back to front. */
    public List<SheetImage> getImages() {
        return Collections.unmodifiableList(images);
    }

    /** Adds a picture (ignored while the sheet is protected). */
    public void addImage(SheetImage image) {
        if (!protection.isProtected()) images.add(image);
    }

    public void removeImage(SheetImage image) {
        if (!protection.isProtected()) images.remove(image);
    }

    /** Read-only view of the charts, back to front. */
    public List<ChartSpec> getCharts() {
        return Collections.unmodifiableList(charts);
    }

    public void addChart(ChartSpec chart) {
        if (!protection.isProtected()) charts.add(chart);
    }

    public void removeChart(ChartSpec chart) {
        if (!protection.isProtected()) charts.remove(chart);
    }

    /** Display text for the cell: formula results are formatted, empty cells give "". */
    @Override
    public Object getValueAt(int row, int col) {
        Object v = getComputedValue(row, col);
        if (v == null) return "";
        if (showFormulas) {
            String f = formulas.get(key(row, col));
            if (f != null) return f;
        }
        if (v instanceof Double) {
            CellStyle st = styles.get(key(row, col));
            if (st != null && st.numFmt != CellFormats.GENERAL) {
                return CellFormats.format((Double) v, st.numFmt, st.decimals);
            }
            return FormulaEngine.formatNumber((Double) v);
        }
        if (v instanceof Boolean) return ((Boolean) v) ? "TRUE" : "FALSE";
        return v.toString();
    }

    public boolean isShowFormulas() {
        return showFormulas;
    }

    /** View-only switch (no table event, so it does not count as an edit). */
    public void setShowFormulas(boolean on) {
        showFormulas = on;
    }

    @Override
    public void setValueAt(Object value, int row, int col) {
        setInput(row, col, value == null ? "" : value.toString());
    }

    // ------------------------------------------------------------------
    // Cell content
    // ------------------------------------------------------------------

    public boolean inBounds(int row, int col) {
        return row >= 0 && col >= 0 && row < rows && col < cols;
    }

    private static String key(int row, int col) {
        return FormulaEngine.cellRef(row, col);
    }

    /** What the user typed: the formula text (with '=') or the literal value. */
    public String getInput(int row, int col) {
        if (!inBounds(row, col)) return "";
        String k = key(row, col);
        String f = formulas.get(k);
        if (f != null) return f;
        String r = rawValues.get(k);
        return r == null ? "" : r;
    }

    /** Stores user input, routing it to the formula or raw-value map, then recalculates. */
    public void setInput(int row, int col, String text) {
        if (!inBounds(row, col) || protection.isProtected()) return;
        String k = key(row, col);
        String t = text == null ? "" : text.trim();
        formulas.remove(k);
        rawValues.remove(k);
        if (t.startsWith("=")) {
            formulas.put(k, t);
        } else if (!t.isEmpty()) {
            Matcher pm = PERCENT_INPUT.matcher(t);
            if (pm.matches()) { // typing 25% stores 0.25 and formats the cell as a percentage
                String num = pm.group(1);
                BigDecimal value = new BigDecimal(num).movePointLeft(2);
                rawValues.put(k, value.compareTo(BigDecimal.ZERO) == 0 ? "0" : value.stripTrailingZeros().toPlainString());
                int dot = num.indexOf('.');
                CellStyle st = styles.get(k);
                if (st == null) {
                    st = new CellStyle();
                    styles.put(k, st);
                }
                if (st.numFmt == CellFormats.GENERAL) {
                    st.numFmt = CellFormats.PERCENT;
                    st.decimals = dot < 0 ? 0 : num.length() - dot - 1;
                }
            } else {
                rawValues.put(k, t);
            }
        }
        recalculate();
    }

    /** Removes the content of a cell but keeps its formatting (like pressing Delete). */
    public void clearContents(int row, int col) {
        setInput(row, col, "");
    }

    /**
     * Evaluated value: Double, String (text or error code), Boolean or null (empty).
     * Typing TRUE or FALSE gives a Boolean, like a real spreadsheet.
     * Formula results are memoised in the cache until the next edit.
     */
    public Object getComputedValue(int row, int col) {
        if (!inBounds(row, col)) return null;
        String k = key(row, col);
        String f = formulas.get(k);
        if (f != null) {
            if (computed.containsKey(k)) return computed.get(k);
            if (!evaluating.add(k)) return FormulaEngine.ERR_GENERIC; // circular reference
            Object result;
            try {
                result = engine.evaluate(f);
            } finally {
                evaluating.remove(k);
            }
            computed.put(k, result);
            return result;
        }
        String raw = rawValues.get(k);
        if (raw == null) return null;
        String trimmed = raw.trim();
        if (trimmed.equalsIgnoreCase("true")) return Boolean.TRUE;
        if (trimmed.equalsIgnoreCase("false")) return Boolean.FALSE;
        Double d = tryParseNumber(raw);
        return d != null ? (Object) d : raw;
    }

    private static Double tryParseNumber(String s) {
        String t = s.trim();
        if (!NUMBER.matcher(t).matches()) return null;
        try {
            return Double.valueOf(t);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public boolean isNumeric(int row, int col) {
        return getComputedValue(row, col) instanceof Double;
    }

    public boolean isFormulaError(int row, int col) {
        if (!inBounds(row, col) || !formulas.containsKey(key(row, col))) return false;
        Object v = getComputedValue(row, col);
        return v instanceof String && FormulaEngine.isErrorCode((String) v);
    }

    // ------------------------------------------------------------------
    // Formatting
    // ------------------------------------------------------------------

    /** @return the cell's style or null when it has default formatting. */
    public CellStyle getStyle(int row, int col) {
        return inBounds(row, col) ? styles.get(key(row, col)) : null;
    }

    /** Applies a mutation to the cell's style (created on demand) and repaints it. */
    public void updateStyle(int row, int col, Consumer<CellStyle> change) {
        if (!inBounds(row, col) || protection.isProtected()) return;
        String k = key(row, col);
        CellStyle st = styles.get(k);
        if (st == null) st = new CellStyle();
        change.accept(st);
        if (st.isDefault()) styles.remove(k);
        else styles.put(k, st);
        if (batchDepth == 0) fireTableCellUpdated(row, col);
        else dirty = true;
    }

    // ------------------------------------------------------------------
    // Recalculation / batching
    // ------------------------------------------------------------------

    /** Invalidates cached results; dependent formulas recompute on next access. */
    private void recalculate() {
        computed.clear();
        dupCache.clear();
        if (batchDepth == 0) fireTableRowsUpdated(0, rows - 1);
        else dirty = true;
    }

    /** Suppresses repaint events until the matching {@link #endBatch()} (used by macros / inserts). */
    public void beginBatch() {
        batchDepth++;
    }

    public void endBatch() {
        if (batchDepth > 0 && --batchDepth == 0 && dirty) {
            dirty = false;
            computed.clear();
            dupCache.clear();
            fireTableRowsUpdated(0, rows - 1);
        }
    }


    // ------------------------------------------------------------------
    // Conditional formatting
    // ------------------------------------------------------------------

    public List<CondRule> getCondRules() {
        return Collections.unmodifiableList(condRules);
    }

    public void addCondRule(CondRule rule) {
        if (protection.isProtected()) return;
        condRules.add(rule);
        dupCache.clear();
    }

    /** Removes every rule that touches the given range. @return how many were removed. */
    public int clearCondRules(int r1, int c1, int r2, int c2) {
        if (protection.isProtected()) return 0;
        int before = condRules.size();
        condRules.removeIf(cr -> cr.r1 <= r2 && cr.r2 >= r1 && cr.c1 <= c2 && cr.c2 >= c1);
        dupCache.clear();
        return before - condRules.size();
    }

    public int clearAllCondRules() {
        if (protection.isProtected()) return 0;
        int n = condRules.size();
        condRules.clear();
        dupCache.clear();
        return n;
    }

    /** The first rule that formats this cell right now, or null. */
    public CondRule matchingRule(int row, int col) {
        if (condRules.isEmpty()) return null;
        for (CondRule rule : condRules) {
            if (rule.covers(row, col) && ruleMatches(rule, row, col)) return rule;
        }
        return null;
    }

    private boolean ruleMatches(CondRule rule, int row, int col) {
        Object v = getComputedValue(row, col);
        if (v == null) return false;
        switch (rule.type) {
            case GREATER:
            case LESS: {
                Double t = tryParseNumber(rule.v1);
                if (!(v instanceof Double) || t == null) return false;
                return rule.type == CondRule.Type.GREATER ? (Double) v > t : (Double) v < t;
            }
            case EQUAL: {
                Double t = tryParseNumber(rule.v1);
                if (v instanceof Double && t != null) return Math.abs((Double) v - t) < 1e-12;
                return String.valueOf(v).trim().equalsIgnoreCase(rule.v1.trim());
            }
            case BETWEEN: {
                Double a = tryParseNumber(rule.v1);
                Double b = tryParseNumber(rule.v2);
                if (!(v instanceof Double) || a == null || b == null) return false;
                double lo = Math.min(a, b);
                double hi = Math.max(a, b);
                return (Double) v >= lo && (Double) v <= hi;
            }
            case CONTAINS:
                return !rule.v1.isEmpty() && String.valueOf(v).toLowerCase().contains(rule.v1.toLowerCase());
            default: { // DUPLICATE
                Map<String, Integer> counts = dupCache.get(rule);
                if (counts == null) {
                    counts = new HashMap<>();
                    for (int r = rule.r1; r <= rule.r2 && r < rows; r++) {
                        for (int c = rule.c1; c <= rule.c2 && c < cols; c++) {
                            Object x = getComputedValue(r, c);
                            if (x != null) counts.merge(String.valueOf(x).toLowerCase(), 1, Integer::sum);
                        }
                    }
                    dupCache.put(rule, counts);
                }
                Integer n = counts.get(String.valueOf(v).toLowerCase());
                return n != null && n > 1;
            }
        }
    }

    // ------------------------------------------------------------------
    // Filter
    // ------------------------------------------------------------------

    public AutoFilter getFilter() {
        return filter;
    }

    public void enableFilter(int headerRow, int firstCol, int lastCol) {
        if (protection.isProtected()) return;
        filter = new AutoFilter(headerRow, firstCol, lastCol);
    }

    public void disableFilter() {
        if (!protection.isProtected()) filter = null;
    }

    /** Sets the visible values of a column ({@code null} = show everything). */
    public void setFilterValues(int col, java.util.Set<String> allowedValues) {
        if (filter == null || protection.isProtected()) return;
        if (allowedValues == null) filter.allowed.remove(col);
        else filter.allowed.put(col, allowedValues);
    }

    public void clearFilterCriteria() {
        if (filter != null && !protection.isProtected()) filter.allowed.clear();
    }

    /** The text the filter list shows for a cell. */
    public String filterKey(int row, int col) {
        String s = String.valueOf(getValueAt(row, col));
        return s.isEmpty() ? AutoFilter.BLANKS : s;
    }

    /** Which rows the filter currently hides (all false without a filter). */
    public boolean[] computeHiddenRows() {
        boolean[] hidden = new boolean[rows];
        if (filter == null || !filter.anyCriteria()) return hidden;
        int last = getUsedRowCount() - 1;
        for (int r = filter.headerRow + 1; r <= last && r < rows; r++) {
            for (Map.Entry<Integer, java.util.Set<String>> e : filter.allowed.entrySet()) {
                if (!e.getValue().contains(filterKey(r, e.getKey()))) {
                    hidden[r] = true;
                    break;
                }
            }
        }
        return hidden;
    }

    /** Last data row the filter covers. */
    public int filterLastRow() {
        return filter == null ? -1 : Math.max(filter.headerRow, getUsedRowCount() - 1);
    }

    // ------------------------------------------------------------------
    // Structure edits: insert / delete rows and columns, sort, duplicates, text to columns
    // ------------------------------------------------------------------

    private void putCell(int r, int c, String input, CellStyle st) {
        String k = key(r, c);
        formulas.remove(k);
        rawValues.remove(k);
        styles.remove(k);
        if (input != null && !input.isEmpty()) {
            if (input.startsWith("=")) formulas.put(k, input);
            else rawValues.put(k, input);
        }
        if (st != null && !st.isDefault()) styles.put(k, st);
    }

    private void structureChange(boolean rowsAxis, int index, int n, boolean insert) {
        int size = rowsAxis ? rows : cols;
        int[] map = new int[size];
        for (int i = 0; i < size; i++) {
            int j;
            if (insert) j = i >= index ? i + n : i;
            else if (i >= index && i < index + n) j = -1;
            else j = i >= index + n ? i - n : i;
            map[i] = j >= size ? -1 : j;
        }
        List<Object[]> cells = new ArrayList<>();
        for (int r = 0; r < rows; r++) {
            for (int c = 0; c < cols; c++) {
                String in = getInput(r, c);
                CellStyle st = getStyle(r, c);
                if (!in.isEmpty() || st != null) cells.add(new Object[]{r, c, in, st});
            }
        }
        formulas.clear();
        rawValues.clear();
        styles.clear();
        for (Object[] e : cells) {
            int r = (Integer) e[0];
            int c = (Integer) e[1];
            int nr = rowsAxis ? map[r] : r;
            int nc = rowsAxis ? c : map[c];
            if (nr < 0 || nc < 0) continue;
            String in = (String) e[2];
            if (in.startsWith("=")) in = FormulaRefs.shift(in, rowsAxis, index, n, insert);
            putCell(nr, nc, in, (CellStyle) e[3]);
        }
        if (!rowsAxis) {
            int[] old = columnWidths.clone();
            java.util.Arrays.fill(columnWidths, DEFAULT_COLUMN_WIDTH);
            for (int i = 0; i < size; i++) if (map[i] >= 0) columnWidths[map[i]] = old[i];
        }
        List<CondRule> remapped = new ArrayList<>();
        for (CondRule cr : condRules) {
            int a = rowsAxis ? cr.r1 : cr.c1;
            int b = rowsAxis ? cr.r2 : cr.c2;
            int lo = Integer.MAX_VALUE;
            int hi = -1;
            for (int i = a; i <= b && i < size; i++) {
                if (map[i] >= 0) {
                    lo = Math.min(lo, map[i]);
                    hi = Math.max(hi, map[i]);
                }
            }
            if (hi < 0) continue;
            remapped.add(rowsAxis ? cr.withRange(lo, cr.c1, hi, cr.c2) : cr.withRange(cr.r1, lo, cr.r2, hi));
        }
        condRules.clear();
        condRules.addAll(remapped);
        if (filter != null) {
            if (rowsAxis) {
                int h = map[filter.headerRow];
                if (h < 0) filter = null;
                else filter.headerRow = h;
            } else {
                int lo = Integer.MAX_VALUE;
                int hi = -1;
                for (int i = filter.firstCol; i <= filter.lastCol && i < size; i++) {
                    if (map[i] >= 0) {
                        lo = Math.min(lo, map[i]);
                        hi = Math.max(hi, map[i]);
                    }
                }
                if (hi < 0) {
                    filter = null;
                } else {
                    Map<Integer, java.util.Set<String>> old = new java.util.LinkedHashMap<>(filter.allowed);
                    filter.allowed.clear();
                    for (Map.Entry<Integer, java.util.Set<String>> e : old.entrySet()) {
                        int nc = e.getKey() < size ? map[e.getKey()] : -1;
                        if (nc >= 0) filter.allowed.put(nc, e.getValue());
                    }
                    filter.firstCol = lo;
                    filter.lastCol = hi;
                }
            }
        }
        recalculate();
    }

    public boolean insertRows(int at, int n) {
        if (protection.isProtected() || n < 1 || at < 0 || at >= rows) return false;
        structureChange(true, at, n, true);
        return true;
    }

    public boolean deleteRows(int at, int n) {
        if (protection.isProtected() || n < 1 || at < 0 || at >= rows) return false;
        structureChange(true, at, Math.min(n, rows - at), false);
        return true;
    }

    public boolean insertColumns(int at, int n) {
        if (protection.isProtected() || n < 1 || at < 0 || at >= cols) return false;
        structureChange(false, at, n, true);
        return true;
    }

    public boolean deleteColumns(int at, int n) {
        if (protection.isProtected() || n < 1 || at < 0 || at >= cols) return false;
        structureChange(false, at, Math.min(n, cols - at), false);
        return true;
    }

    private static int sortRank(Object v) {
        if (v == null) return 4;
        if (v instanceof Double) return 0;
        if (v instanceof Boolean) return 2;
        return FormulaEngine.isErrorCode((String) v) ? 3 : 1;
    }

    private static int compareForSort(Object a, Object b, boolean asc) {
        int ra = sortRank(a);
        int rb = sortRank(b);
        if (ra == 4 || rb == 4) return ra == rb ? 0 : (ra == 4 ? 1 : -1); // blanks always last
        int cmp;
        if (ra != rb) cmp = Integer.compare(ra, rb);
        else if (ra == 0) cmp = Double.compare((Double) a, (Double) b);
        else if (ra == 2) cmp = Boolean.compare((Boolean) a, (Boolean) b);
        else cmp = String.CASE_INSENSITIVE_ORDER.compare((String) a, (String) b);
        return asc ? cmp : -cmp;
    }

    /**
     * Sorts the rows r1..r2 (columns c1..c2 move together) by the values in keyCol.
     * Formulas that refer to their own row keep pointing at it. Blank cells always sort last.
     */
    public boolean sortRows(int r1, int r2, int c1, int c2, int keyCol, boolean asc) {
        if (protection.isProtected() || r2 <= r1 || !inBounds(r1, c1) || !inBounds(r2, c2)) return false;
        int n = r2 - r1 + 1;
        int w = c2 - c1 + 1;
        Integer[] order = new Integer[n];
        Object[] keys = new Object[n];
        String[][] in = new String[n][w];
        CellStyle[][] st = new CellStyle[n][w];
        for (int i = 0; i < n; i++) {
            order[i] = i;
            keys[i] = getComputedValue(r1 + i, keyCol);
            for (int c = 0; c < w; c++) {
                in[i][c] = getInput(r1 + i, c1 + c);
                CellStyle s0 = getStyle(r1 + i, c1 + c);
                st[i][c] = s0 == null ? null : s0.copy();
            }
        }
        java.util.Arrays.sort(order, (a, b) -> compareForSort(keys[a], keys[b], asc)); // stable
        for (int k = 0; k < n; k++) {
            int src = order[k];
            for (int c = 0; c < w; c++) {
                putCell(r1 + k, c1 + c, FormulaRefs.moveRow(in[src][c], r1 + src, r1 + k), st[src][c]);
            }
        }
        recalculate();
        return true;
    }

    /** Removes later rows that repeat an earlier row's values in c1..c2. @return rows removed, -1 if blocked. */
    public int removeDuplicateRows(int r1, int r2, int c1, int c2) {
        if (protection.isProtected() || r2 < r1) return -1;
        int n = r2 - r1 + 1;
        int w = c2 - c1 + 1;
        java.util.Set<String> seen = new java.util.HashSet<>();
        List<Integer> keep = new ArrayList<>();
        for (int i = 0; i < n; i++) {
            StringBuilder sb = new StringBuilder();
            boolean blank = true;
            for (int c = c1; c <= c2; c++) {
                String t = getInput(r1 + i, c);
                if (!t.isEmpty()) blank = false;
                sb.append(t.toLowerCase()).append('\u0001');
            }
            if (blank || seen.add(sb.toString())) keep.add(i);
        }
        int removed = n - keep.size();
        if (removed == 0) return 0;
        String[][] in = new String[n][w];
        CellStyle[][] st = new CellStyle[n][w];
        for (int i = 0; i < n; i++) {
            for (int c = 0; c < w; c++) {
                in[i][c] = getInput(r1 + i, c1 + c);
                CellStyle s0 = getStyle(r1 + i, c1 + c);
                st[i][c] = s0 == null ? null : s0.copy();
            }
        }
        for (int k = 0; k < keep.size(); k++) {
            int src = keep.get(k);
            for (int c = 0; c < w; c++) {
                putCell(r1 + k, c1 + c, FormulaRefs.moveRow(in[src][c], r1 + src, r1 + k), st[src][c]);
            }
        }
        for (int r = r1 + keep.size(); r <= r2; r++) for (int c = c1; c <= c2; c++) putCell(r, c, null, null);
        recalculate();
        return removed;
    }

    /** The block of filled cells around (r, c), like Excel's Ctrl+A on a table: {r1, c1, r2, c2}. */
    public int[] currentRegion(int r, int c) {
        int r1 = r;
        int r2 = r;
        int c1 = c;
        int c2 = c;
        boolean grew = true;
        while (grew) {
            grew = false;
            int ca = Math.max(0, c1 - 1);
            int cb = Math.min(cols - 1, c2 + 1);
            int ra = Math.max(0, r1 - 1);
            int rb = Math.min(rows - 1, r2 + 1);
            if (r1 > 0 && rangeHasData(r1 - 1, ca, r1 - 1, cb)) {
                r1--;
                grew = true;
            }
            if (r2 < rows - 1 && rangeHasData(r2 + 1, ca, r2 + 1, cb)) {
                r2++;
                grew = true;
            }
            if (c1 > 0 && rangeHasData(ra, c1 - 1, rb, c1 - 1)) {
                c1--;
                grew = true;
            }
            if (c2 < cols - 1 && rangeHasData(ra, c2 + 1, rb, c2 + 1)) {
                c2++;
                grew = true;
            }
        }
        return new int[]{r1, c1, r2, c2};
    }

    private boolean rangeHasData(int r1, int c1, int r2, int c2) {
        for (int r = r1; r <= r2; r++) for (int c = c1; c <= c2; c++) if (!getInput(r, c).isEmpty()) return true;
        return false;
    }

    /** Splits the text in column {@code col} at the delimiter into the columns to its right. @return columns used, -1 if blocked. */
    public int textToColumns(int r1, int r2, int col, String delimiter) {
        if (protection.isProtected() || delimiter == null || delimiter.isEmpty()) return -1;
        int maxParts = 1;
        String q = Pattern.quote(delimiter);
        for (int r = r1; r <= r2; r++) {
            String in = getInput(r, col);
            if (in.isEmpty() || in.startsWith("=")) continue;
            String[] parts = in.split(q, -1);
            maxParts = Math.max(maxParts, parts.length);
            for (int k = 0; k < parts.length && col + k < cols; k++) {
                CellStyle st = getStyle(r, col + k);
                putCell(r, col + k, parts[k].trim(), st == null ? null : st.copy());
            }
        }
        recalculate();
        return maxParts;
    }

    /** True when any cell in the range holds something. */
    public boolean hasDataIn(int r1, int c1, int r2, int c2) {
        return rangeHasData(Math.max(0, r1), Math.max(0, c1), Math.min(rows - 1, r2), Math.min(cols - 1, c2));
    }

    /** Raw writes for tools that build whole blocks (pivot tables, summaries); recalculates once at the end. */
    public void setInputs(int r, int c, String text, CellStyle style) {
        if (protection.isProtected() || !inBounds(r, c)) return;
        putCell(r, c, text == null ? "" : text.trim(), style);
    }

    public void refreshAll() {
        recalculate();
    }

    public Set<String> distinctFilterKeys(int col, int fromRow, int toRow) {
        Set<String> keys = new LinkedHashSet<>();
        for (int r = fromRow; r <= toRow && r < rows; r++) keys.add(filterKey(r, col));
        return keys;
    }

    // ------------------------------------------------------------------
    // Used range (for CSV export)
    // ------------------------------------------------------------------

    public int getUsedRowCount() {
        int last = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!getInput(r, c).isEmpty()) last = r + 1;
        return last;
    }

    public int getUsedColumnCount() {
        int last = 0;
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                if (!getInput(r, c).isEmpty() && c + 1 > last) last = c + 1;
        return last;
    }
}
