package Util;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentForm;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCProperty;

public class ExcelUtil {

    private static final String VML_REL =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

    private static final String VML_NS =
            "urn:schemas-microsoft-com:vml";

    private static final String EXCEL_NS =
            "urn:schemas-microsoft-com:office:excel";

    private static final String CTRL_PROP_REL =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";

    private static final String REL_NS =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static final String MAIN_NS =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static final String CTRL_PROP_NS =
            "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    private static final int DEFAULT_CHECKBOX_SCAN_ROWS = 8;

    private static final Set<String> CHECKBOX_PROPERTIES =
            new LinkedHashSet<String>(Arrays.asList(
                    "a2Manufacturer",
                    "a2ProjectType",
                    "a2NoOfCavityNew",
                    "a2TypeofRunner",
                    "a2TypeOfSubRunner",
                    "a2GatetypeNew",
                    "a2HRS_OEM",
                    "a2TypeofmoldingmachineNew",
                    "a2StdBoughtoutItemMouldBase",
                    "a2EjectorPin",
                    "a2EjectorPinType",
                    "a2EjectorPinMake",
                    "a2NumberingOrCoringPin",
                    "a2CenteringSleeve",
                    "a2LocationCoolingConnectors",
                    "a2AirVentMoldexAnalysis"
            ));

    private static final Set<String> SINGLE_ROW_CHECKBOX_PROPERTIES =
            new LinkedHashSet<String>(Arrays.asList(
                    "a2Manufacturer",
                    "a2ProjectType",
                    "a2NoOfCavityNew",
                    "a2HRS_OEM"
            ));

    public static void processSheet(
            Sheet sheet,
            TCComponentItemRevision itemRev,
            Map<String, String> mapping) {

        DataFormatter formatter = new DataFormatter();
        VmlCheckboxEditor checkboxEditor = VmlCheckboxEditor.open(sheet);

        try {
            for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {

                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }

                for (Cell cell : row) {

                    String label = formatter.formatCellValue(cell).trim();
                    if (label.isEmpty()) {
                        continue;
                    }

                    String prop = getMappedProperty(label, mapping);
                    if (prop == null) {
                        continue;
                    }

                    System.out.println("Processing Label: " + label + " -> " + prop);

                    int nextLabelCol = findNextLabelColumnOnRow(
                            row, cell.getColumnIndex(), mapping, formatter);

                    if ("MULTI_PLM_DRAWING_REV".equals(prop)) {
                        Cell valueCell = findSingleValueCell(
                                sheet, rowNum, cell, nextLabelCol, formatter, mapping);

                        if (valueCell != null) {
                            valueCell.setCellValue(getCombinedPlmDrawingRevision(itemRev));
                        }
                        continue;
                    }

                    Set<String> selectedValues = getSelectedValues(itemRev, prop);
                    System.out.println("TC value for " + prop + " = " + selectedValues);

                    boolean forceCheckbox = isCheckboxProperty(prop);

                    int blockEndRow = forceCheckbox
                            ? findCheckboxBlockEndRow(sheet, rowNum, cell.getColumnIndex(), formatter, mapping)
                            : findBlockEndRow(sheet, rowNum, nextLabelCol, mapping, formatter);

                    if (SINGLE_ROW_CHECKBOX_PROPERTIES.contains(prop)) {
                        blockEndRow = rowNum;
                    }

                    boolean checkboxHandled = checkboxEditor != null
                            && checkboxEditor.updateBlock(
                                    sheet,
                                    rowNum,
                                    cell.getColumnIndex(),
                                    nextLabelCol,
                                    blockEndRow,
                                    prop,
                                    selectedValues,
                                    formatter);

                    if (!checkboxHandled && !forceCheckbox) {
                        Cell valueCell = findSingleValueCell(
                                sheet, rowNum, cell, nextLabelCol, formatter, mapping);

                        if (valueCell != null) {
                            valueCell.setCellValue(joinValues(selectedValues));
                        }
                    }
                }
            }

            if (checkboxEditor != null) {
                checkboxEditor.save();
            }

        } catch (Exception e) {
            throw new RuntimeException("Failed while processing template", e);
        }
    }

    public static void processTableBlocks(
            Sheet sheet,
            TCComponentItemRevision itemRev,
            Map<String, String> tableMapping) {

        DataFormatter formatter = new DataFormatter();

        for (Map.Entry<String, String> entry : tableMapping.entrySet()) {

            String[] parts = entry.getKey().split("\\|", 3);
            if (parts.length != 3) {
                continue;
            }

            String section = parts[0].trim();
            String rowHeader = parts[1].trim();
            String columnHeader = parts[2].trim();
            String prop = entry.getValue();

            Cell targetCell = findTableCell(sheet, section, rowHeader, columnHeader, formatter);
            if (targetCell == null) {
                System.out.println("Table cell not found for: " + entry.getKey());
                continue;
            }

            String value = joinValues(getSelectedValues(itemRev, prop));
            targetCell.setCellValue(value);

            System.out.println("Filled table cell: " + entry.getKey() + " -> " + value);
        }
    }

    private static boolean isCheckboxProperty(String prop) {
        return CHECKBOX_PROPERTIES.contains(prop);
    }

    private static int findBlockEndRow(
            Sheet sheet,
            int labelRowNum,
            int nextLabelCol,
            Map<String, String> mapping,
            DataFormatter formatter) {

        int endRow = labelRowNum;

        for (int rowNum = labelRowNum + 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }

            if (rowContainsAnotherMappedLabel(row, mapping, formatter)
                    || rowContainsContentInColumnRange(row, nextLabelCol, formatter)) {
                break;
            }

            endRow = rowNum;
        }

        return endRow;
    }

    private static int findCheckboxBlockEndRow(
            Sheet sheet,
            int labelRowNum,
            int labelCol,
            DataFormatter formatter,
            Map<String, String> mapping) {

        int endRow = Math.min(sheet.getLastRowNum(), labelRowNum + DEFAULT_CHECKBOX_SCAN_ROWS);

        for (int rowNum = labelRowNum + 1; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }

            for (Cell cell : row) {
                if (cell.getColumnIndex() <= labelCol) {
                    String text = formatter.formatCellValue(cell).trim();
                    if (!text.isEmpty() && getMappedProperty(text, mapping) != null) {
                        return rowNum - 1;
                    }
                }
            }
        }

        return endRow;
    }

    private static Cell findSingleValueCell(
            Sheet sheet,
            int rowNum,
            Cell labelCell,
            int nextLabelCol,
            DataFormatter formatter,
            Map<String, String> mapping) {

        Row row = sheet.getRow(rowNum);
        if (row == null) {
            return null;
        }

        int startCol = getStartColumnAfterLabel(sheet, rowNum, labelCell.getColumnIndex());
        int searchEnd = nextLabelCol > -1
                ? nextLabelCol
                : Math.max(row.getLastCellNum(), startCol + 6);

        for (int col = startCol; col < searchEnd; col++) {

            Cell cell = row.getCell(col);
            if (cell == null) {
                return row.createCell(col);
            }

            String text = formatter.formatCellValue(cell).trim();

            if (text.isEmpty()) {
                return cell;
            }

            if (getMappedProperty(text, mapping) != null) {
                return null;
            }
        }

        return row.createCell(startCol);
    }

    private static int findNextLabelColumnOnRow(
            Row row,
            int currentLabelCol,
            Map<String, String> mapping,
            DataFormatter formatter) {

        for (Cell cell : row) {
            if (cell.getColumnIndex() <= currentLabelCol) {
                continue;
            }

            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty() && getMappedProperty(text, mapping) != null) {
                return cell.getColumnIndex();
            }
        }

        return -1;
    }

    private static int getStartColumnAfterLabel(Sheet sheet, int rowNum, int labelCol) {

        int startCol = labelCol + 1;

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.isInRange(rowNum, labelCol)) {
                startCol = range.getLastColumn() + 1;
                break;
            }
        }

        return startCol;
    }

    private static boolean rowContainsAnotherMappedLabel(
            Row row,
            Map<String, String> mapping,
            DataFormatter formatter) {

        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty() && getMappedProperty(text, mapping) != null) {
                return true;
            }
        }

        return false;
    }

    private static boolean rowContainsContentInColumnRange(
            Row row,
            int startCol,
            DataFormatter formatter) {

        if (startCol < 0) {
            return false;
        }

        for (int col = startCol; col < row.getLastCellNum(); col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }

            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty()) {
                return true;
            }
        }

        return false;
    }

    private static Set<String> getSelectedValues(TCComponentItemRevision itemRev, String prop) {

        Set<String> values = new LinkedHashSet<String>();
        collectValues(itemRev, prop, values);

        if (!values.isEmpty()) {
            return values;
        }

        try {
            TCComponent[] forms = itemRev.getRelatedComponents("IMAN_specification");
            for (TCComponent comp : forms) {
                if (comp instanceof TCComponentForm) {
                    collectValues((TCComponentForm) comp, prop, values);
                    if (!values.isEmpty()) {
                        return values;
                    }
                }
            }
        } catch (Exception ignored) {
        }

        return values;
    }

    private static void collectValues(TCComponent component, String prop, Set<String> values) {

        if (component == null) {
            return;
        }

        try {
            TCProperty tcProp = component.getTCProperty(prop);
            if (tcProp != null) {
                addValues(tcProp.getDisplayableValues(), values);

                String single = tcProp.getDisplayableValue();
                if (single != null && !single.trim().isEmpty()) {
                    addSplitValues(single, values);
                }
            }
        } catch (Exception ignored) {
        }

        try {
            addSplitValues(component.getProperty(prop), values);
        } catch (Exception ignored) {
        }
    }

    private static void addValues(Object rawValues, Set<String> values) {

        if (rawValues == null) {
            return;
        }

        if (rawValues instanceof String[]) {
            String[] arr = (String[]) rawValues;
            for (String value : arr) {
                addSplitValues(value, values);
            }
            return;
        }

        if (rawValues instanceof List<?>) {
            List<?> list = (List<?>) rawValues;
            for (Object value : list) {
                if (value != null) {
                    addSplitValues(String.valueOf(value), values);
                }
            }
            return;
        }

        addSplitValues(String.valueOf(rawValues), values);
    }

    private static void addSplitValues(String raw, Set<String> values) {

        if (raw == null || raw.trim().isEmpty()) {
            return;
        }

        String[] parts = raw.split("[,;\\n\\r]+");
        for (String part : parts) {
            if (!part.trim().isEmpty()) {
                values.add(part.trim());
            }
        }
    }

    private static String joinValues(Set<String> values) {

        if (values.isEmpty()) {
            return "";
        }

        StringBuilder sb = new StringBuilder();
        for (String value : values) {
            if (sb.length() > 0) {
                sb.append(", ");
            }
            sb.append(value);
        }
        return sb.toString();
    }

    private static boolean matchesAny(String optionText, Set<String> selectedValues) {

        String normalizedOption = canonicalizeValue(optionText);
        if (normalizedOption.isEmpty()) {
            return false;
        }

        for (String selected : selectedValues) {
            if (matchesOption(normalizedOption, selected)) {
                return true;
            }
        }

        return false;
    }

    private static boolean matchesOption(String normalizedOption, String selectedValue) {

        String normalizedSelected = canonicalizeValue(selectedValue);
        if (normalizedSelected.isEmpty()) {
            return false;
        }

        return normalizedOption.equals(normalizedSelected);
    }

    private static String canonicalizeValue(String value) {

        String v = normalize(value);

        if ("submarinegate".equals(v) || "submarine".equals(v)) {
            return "submarine";
        }

        if ("pinpoint".equals(v) || "pinpointgate".equals(v) || "point".equals(v)) {
            return "point";
        }

        if ("inhouse".equals(v) || "inhouseproject".equals(v)) {
            return "inhouse";
        }

        if ("outsource".equals(v) || "outsourced".equals(v)) {
            return "outsourced";
        }

        return v;
    }

    private static String normalize(String text) {

        if (text == null) {
            return "";
        }

        return text.replaceAll("[^a-zA-Z0-9]", "").toLowerCase();
    }

    private static String getMappedProperty(String label, Map<String, String> mapping) {

        String cleanLabel = label.replaceAll("\\s+", "").toLowerCase();

        for (Map.Entry<String, String> entry : mapping.entrySet()) {
            String key = entry.getKey().replaceAll("\\s+", "").toLowerCase();
            if (key.equals(cleanLabel)) {
                return entry.getValue();
            }
        }

        return null;
    }

    private static String getCombinedPlmDrawingRevision(TCComponentItemRevision itemRev) {

        String itemId = safeGetProperty(itemRev, "item_id");
        String drawingNo = safeGetProperty(itemRev, "a2DrawingNo");

        String revNo = safeGetProperty(itemRev, "a2RevNoSMP");
        if (revNo.isEmpty()) {
            revNo = safeGetProperty(itemRev, "a2RevNoDPP");
        }

        StringBuilder sb = new StringBuilder();

        if (!itemId.isEmpty()) {
            sb.append(itemId);
        }

        if (!drawingNo.isEmpty()) {
            if (sb.length() > 0) {
                sb.append(" / ");
            }
            sb.append(drawingNo);
        }

        if (!revNo.isEmpty()) {
            if (sb.length() > 0) {
                sb.append(" / ");
            }
            sb.append(revNo);
        }

        return sb.toString();
    }

    private static String safeGetProperty(TCComponent component, String prop) {
        try {
            String value = component.getProperty(prop);
            return value == null ? "" : value.trim();
        } catch (Exception e) {
            return "";
        }
    }

    private static boolean isOtherOption(String optionText) {
        return "other".equals(normalize(optionText));
    }

    private static Set<String> findUnmatchedValues(Set<String> selectedValues, List<String> optionTexts) {

        Set<String> unmatched = new LinkedHashSet<String>();

        for (String selected : selectedValues) {
            boolean matched = false;

            for (String optionText : optionTexts) {
                if (isOtherOption(optionText)) {
                    continue;
                }

                String normalizedOption = canonicalizeValue(optionText);
                if (matchesOption(normalizedOption, selected)) {
                    matched = true;
                    break;
                }
            }

            if (!matched) {
                unmatched.add(selected);
            }
        }

        return unmatched;
    }

    private static void writeOtherValue(
            Sheet sheet,
            CheckboxShape otherShape,
            int nextLabelCol,
            Set<String> unmatchedValues,
            DataFormatter formatter) {

        Cell valueCell = findOtherValueCell(sheet, otherShape, nextLabelCol, formatter);
        if (valueCell == null) {
            return;
        }

        valueCell.setCellValue(joinValues(unmatchedValues));
    }

    private static Cell findOtherValueCell(
            Sheet sheet,
            CheckboxShape otherShape,
            int nextLabelCol,
            DataFormatter formatter) {

        int[] rowCandidates = new int[] {
                otherShape.row,
                otherShape.row + 1,
                otherShape.row - 1
        };

        for (int rowIndex : rowCandidates) {
            if (rowIndex < 0) {
                continue;
            }

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }

            int endCol = nextLabelCol > -1 ? nextLabelCol : row.getLastCellNum();

            for (int col = otherShape.col; col < endCol; col++) {
                String text = getCellText(sheet, rowIndex, col, formatter);
                if (!isOtherOption(text)) {
                    continue;
                }

                int valueCol = getStartColumnAfterLabel(sheet, rowIndex, col);
                if (nextLabelCol > -1 && valueCol >= nextLabelCol) {
                    return null;
                }

                Cell valueCell = row.getCell(valueCol);
                if (valueCell == null) {
                    valueCell = row.createCell(valueCol);
                }

                return valueCell;
            }
        }

        return null;
    }

    private static Cell findTableCell(
            Sheet sheet,
            String section,
            String rowHeader,
            String columnHeader,
            DataFormatter formatter) {

        int sectionRow = findRowContainingText(sheet, section, formatter);
        if (sectionRow < 0) {
            return null;
        }

        int headerRowNum = -1;
        int rowLabelCol = -1;
        int targetCol = -1;
        int targetRow = -1;

        for (int r = sectionRow + 1; r <= Math.min(sheet.getLastRowNum(), sectionRow + 20); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (normalize(text).equals(normalize("Description"))) {
                    headerRowNum = r;
                    rowLabelCol = cell.getColumnIndex();
                    break;
                }
            }

            if (headerRowNum >= 0) {
                break;
            }
        }

        if (headerRowNum < 0) {
            return null;
        }

        Row headerRow = sheet.getRow(headerRowNum);
        for (Cell cell : headerRow) {
            String text = formatter.formatCellValue(cell).trim();
            if (normalize(text).equals(normalize(columnHeader))) {
                targetCol = cell.getColumnIndex();
                break;
            }
        }

        if (targetCol < 0) {
            return null;
        }

        for (int r = headerRowNum + 1; r <= Math.min(sheet.getLastRowNum(), headerRowNum + 30); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            Cell rowCell = row.getCell(rowLabelCol);
            if (rowCell == null) {
                continue;
            }

            String text = formatter.formatCellValue(rowCell).trim();
            if (normalize(text).equals(normalize(rowHeader))) {
                targetRow = r;
                break;
            }
        }

        if (targetRow < 0) {
            return null;
        }

        Row row = sheet.getRow(targetRow);
        Cell cell = row.getCell(targetCol);
        if (cell == null) {
            cell = row.createCell(targetCol);
        }

        return cell;
    }

    private static int findRowContainingText(
            Sheet sheet,
            String expectedText,
            DataFormatter formatter) {

        String expected = normalize(expectedText);

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (normalize(text).equals(expected)) {
                    return r;
                }
            }
        }

        return -1;
    }

    private static String getCellText(
            Sheet sheet,
            int rowIndex,
            int colIndex,
            DataFormatter formatter) {

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }

        Cell cell = row.getCell(colIndex);
        String text = cell == null ? "" : formatter.formatCellValue(cell).trim();
        if (!text.isEmpty()) {
            return text;
        }

        Cell mergedCell = getMergedRegionPrimaryCell(sheet, rowIndex, colIndex);
        if (mergedCell == null) {
            return "";
        }

        return formatter.formatCellValue(mergedCell).trim();
    }

    private static Cell getMergedRegionPrimaryCell(Sheet sheet, int rowIndex, int colIndex) {

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (!range.isInRange(rowIndex, colIndex)) {
                continue;
            }

            Row firstRow = sheet.getRow(range.getFirstRow());
            if (firstRow == null) {
                return null;
            }

            return firstRow.getCell(range.getFirstColumn());
        }

        return null;
    }

    private static final class VmlCheckboxEditor {

        private final PackagePart vmlPart;
        private final Document vmlDocument;
        private final List<CheckboxShape> shapes;
        private final List<CtrlPropDoc> ctrlPropDocs;

        private VmlCheckboxEditor(
                PackagePart vmlPart,
                Document vmlDocument,
                List<CheckboxShape> shapes,
                List<CtrlPropDoc> ctrlPropDocs) {
            this.vmlPart = vmlPart;
            this.vmlDocument = vmlDocument;
            this.shapes = shapes;
            this.ctrlPropDocs = ctrlPropDocs;
        }

        static VmlCheckboxEditor open(Sheet sheet) {

            if (!(sheet instanceof XSSFSheet)) {
                return null;
            }

            try {
                XSSFSheet xssfSheet = (XSSFSheet) sheet;
                PackagePart sheetPart = xssfSheet.getPackagePart();

                for (PackageRelationship rel : sheetPart.getRelationshipsByType(VML_REL)) {
                    PackagePart vmlPart = sheetPart.getRelatedPart(rel);

                    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                    factory.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
                    factory.setNamespaceAware(true);

                    DocumentBuilder builder = factory.newDocumentBuilder();

                    Document vmlDocument;
                    try (InputStream is = vmlPart.getInputStream()) {
                        vmlDocument = builder.parse(is);
                    }

                    Map<String, CtrlPropDoc> ctrlPropByShapeId =
                            loadCtrlProps(sheetPart, builder);

                    return new VmlCheckboxEditor(
                            vmlPart,
                            vmlDocument,
                            parseShapes(vmlDocument, ctrlPropByShapeId),
                            new ArrayList<CtrlPropDoc>(ctrlPropByShapeId.values()));
                }

            } catch (Exception e) {
                throw new RuntimeException("Failed to open checkbox editor", e);
            }

            return null;
        }

        boolean updateBlock(
                Sheet sheet,
                int labelRowNum,
                int labelCol,
                int nextLabelCol,
                int blockEndRow,
                String prop,
                Set<String> selectedValues,
                DataFormatter formatter) {

            int startCol = labelCol;
            boolean foundCheckbox = false;

            List<CheckboxShape> blockShapes = getBlockShapes(labelRowNum, nextLabelCol, blockEndRow, startCol);
            Map<CheckboxShape, String> optionTextByShape = new HashMap<CheckboxShape, String>();
            List<String> optionTexts = new ArrayList<String>();
            CheckboxShape otherShape = null;

            for (CheckboxShape shape : blockShapes) {

                String optionText = readOptionText(sheet, shape, blockShapes, formatter, prop);
                if (optionText == null || optionText.trim().isEmpty()) {
                    continue;
                }

                foundCheckbox = true;
                optionTextByShape.put(shape, optionText);
                optionTexts.add(optionText);

                if (isOtherOption(optionText)) {
                    otherShape = shape;
                }
            }

            Set<String> unmatchedValues = findUnmatchedValues(selectedValues, optionTexts);

            for (Map.Entry<CheckboxShape, String> entry : optionTextByShape.entrySet()) {
                CheckboxShape shape = entry.getKey();
                String optionText = entry.getValue();

                boolean checked;
                if (isOtherOption(optionText)) {
                    checked = !unmatchedValues.isEmpty();
                } else {
                    checked = matchesAny(optionText, selectedValues);
                }

                setChecked(shape.shapeElement, checked);
                setChecked(shape.ctrlPropDoc, checked);

                System.out.println("Checkbox option = " + optionText + " | checked = " + checked);
            }

            if (otherShape != null) {
                writeOtherValue(sheet, otherShape, nextLabelCol, unmatchedValues, formatter);
            }

            return foundCheckbox;
        }

        private List<CheckboxShape> getBlockShapes(
                int labelRowNum,
                int nextLabelCol,
                int blockEndRow,
                int startCol) {

            List<CheckboxShape> result = new ArrayList<CheckboxShape>();

            for (CheckboxShape shape : shapes) {
                if (shape.row < labelRowNum || shape.row > blockEndRow) {
                    continue;
                }

                if (shape.col < startCol) {
                    continue;
                }

                if (nextLabelCol > -1 && shape.col >= nextLabelCol) {
                    continue;
                }

                result.add(shape);
            }

            return result;
        }

        void save() {

            try (OutputStream os = vmlPart.getOutputStream()) {
                TransformerFactory tf = TransformerFactory.newInstance();
                tf.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);

                Transformer transformer = tf.newTransformer();
                transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
                transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");

                transformer.transform(new DOMSource(vmlDocument), new StreamResult(os));

            } catch (Exception e) {
                throw new RuntimeException("Failed to save VML checkbox data", e);
            }

            for (CtrlPropDoc ctrlPropDoc : ctrlPropDocs) {
                try (OutputStream os = ctrlPropDoc.part.getOutputStream()) {
                    TransformerFactory tf = TransformerFactory.newInstance();
                    tf.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);

                    Transformer transformer = tf.newTransformer();
                    transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
                    transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");

                    transformer.transform(
                            new DOMSource(ctrlPropDoc.document),
                            new StreamResult(os));

                } catch (Exception e) {
                    throw new RuntimeException("Failed to save ctrlProps checkbox data", e);
                }
            }
        }

        private static Map<String, CtrlPropDoc> loadCtrlProps(
                PackagePart sheetPart,
                DocumentBuilder builder) throws Exception {

            Map<String, CtrlPropDoc> result = new HashMap<String, CtrlPropDoc>();

            Document sheetDocument;
            try (InputStream is = sheetPart.getInputStream()) {
                sheetDocument = builder.parse(is);
            }

            NodeList controlNodes = sheetDocument.getElementsByTagNameNS(MAIN_NS, "control");

            for (int i = 0; i < controlNodes.getLength(); i++) {
                Element control = (Element) controlNodes.item(i);

                String shapeId = control.getAttribute("shapeId");
                String relId = control.getAttributeNS(REL_NS, "id");

                if (shapeId == null || shapeId.trim().isEmpty()
                        || relId == null || relId.trim().isEmpty()) {
                    continue;
                }

                PackageRelationship rel = sheetPart.getRelationship(relId);
                if (rel == null || !CTRL_PROP_REL.equals(rel.getRelationshipType())) {
                    continue;
                }

                PackagePart ctrlPropPart = sheetPart.getRelatedPart(rel);

                Document ctrlPropDocument;
                try (InputStream is = ctrlPropPart.getInputStream()) {
                    ctrlPropDocument = builder.parse(is);
                }

                result.put(shapeId, new CtrlPropDoc(ctrlPropPart, ctrlPropDocument));
            }

            return result;
        }

        private static List<CheckboxShape> parseShapes(
                Document vmlDocument,
                Map<String, CtrlPropDoc> ctrlPropByShapeId) {

            List<CheckboxShape> result = new ArrayList<CheckboxShape>();
            NodeList shapeNodes = vmlDocument.getElementsByTagNameNS(VML_NS, "shape");

            for (int i = 0; i < shapeNodes.getLength(); i++) {
                Element shape = (Element) shapeNodes.item(i);

                Element clientData = firstChild(shape, EXCEL_NS, "ClientData");
                if (clientData == null) {
                    continue;
                }

                String objectType = clientData.getAttribute("ObjectType");
                if (!"Checkbox".equalsIgnoreCase(objectType)) {
                    continue;
                }

                Element anchor = firstChild(clientData, EXCEL_NS, "Anchor");
                if (anchor == null) {
                    continue;
                }

                int[] pos = parseAnchor(anchor.getTextContent());
                if (pos == null) {
                    continue;
                }

                String shapeId = parseShapeId(shape.getAttribute("id"));

                result.add(new CheckboxShape(
                        pos[0],
                        pos[1],
                        shape,
                        ctrlPropByShapeId.get(shapeId)));
            }

            return result;
        }

        private static int[] parseAnchor(String anchorText) {

            if (anchorText == null) {
                return null;
            }

            String[] parts = anchorText.trim().split("\\s*,\\s*");
            if (parts.length < 4) {
                return null;
            }

            try {
                int col = Integer.parseInt(parts[0].trim());
                int row = Integer.parseInt(parts[2].trim());
                return new int[] { col, row };
            } catch (Exception e) {
                return null;
            }
        }

        private static String parseShapeId(String rawShapeId) {

            if (rawShapeId == null) {
                return null;
            }

            int index = rawShapeId.lastIndexOf('s');
            if (index >= 0 && index + 1 < rawShapeId.length()) {
                return rawShapeId.substring(index + 1);
            }

            return rawShapeId;
        }

        private static String readOptionText(
                Sheet sheet,
                CheckboxShape currentShape,
                List<CheckboxShape> blockShapes,
                DataFormatter formatter,
                String prop) {

            int nextCheckboxCol = findNextCheckboxColumnSameRow(currentShape, blockShapes);

            if ("a2Manufacturer".equals(prop) || "a2ProjectType".equals(prop)) {
                return readManufacturerProjectTypeText(
                        sheet,
                        currentShape.row,
                        currentShape.col,
                        nextCheckboxCol,
                        formatter,
                        prop);
            }

            if ("a2HRS_OEM".equals(prop) || "a2NoOfCavityNew".equals(prop)) {
                return readCompactCheckboxText(
                        sheet,
                        currentShape.row,
                        currentShape.col,
                        nextCheckboxCol,
                        formatter);
            }

            String text = readBoundedText(sheet, currentShape.row, currentShape.col, nextCheckboxCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readBoundedText(sheet, currentShape.row + 1, currentShape.col, nextCheckboxCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readBoundedText(sheet, currentShape.row - 1, currentShape.col, nextCheckboxCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            return "";
        }

        private static String readManufacturerProjectTypeText(
                Sheet sheet,
                int rowIndex,
                int checkboxCol,
                int nextCheckboxCol,
                DataFormatter formatter,
                String prop) {

            String sameCellText = getCellText(sheet, rowIndex, checkboxCol, formatter).trim();

            if (!sameCellText.isEmpty() && !isHeaderTextForCheckbox(prop, sameCellText)) {
                return sameCellText;
            }

            int limit = nextCheckboxCol == Integer.MAX_VALUE
                    ? checkboxCol + 3
                    : nextCheckboxCol;

            for (int col = checkboxCol + 1; col < limit; col++) {
                String text = getCellText(sheet, rowIndex, col, formatter).trim();
                if (!text.isEmpty()) {
                    return text;
                }
            }

            return "";
        }

        private static boolean isHeaderTextForCheckbox(String prop, String text) {

            String normalized = normalize(text);

            if ("a2Manufacturer".equals(prop)) {
                return "manufacturer".equals(normalized);
            }

            if ("a2ProjectType".equals(prop)) {
                return "projecttype".equals(normalized);
            }

            return false;
        }


        private static String readCompactCheckboxText(
                Sheet sheet,
                int rowIndex,
                int startCol,
                int endCol,
                DataFormatter formatter) {

            String text = readNearestRightText(sheet, rowIndex, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readNearestRightText(sheet, rowIndex + 1, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readNearestRightText(sheet, rowIndex - 1, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readBoundedText(sheet, rowIndex, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readBoundedText(sheet, rowIndex + 1, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            text = readBoundedText(sheet, rowIndex - 1, startCol, endCol, formatter);
            if (!text.isEmpty()) {
                return text;
            }

            return "";
        }

        private static int findNextCheckboxColumnSameRow(
                CheckboxShape currentShape,
                List<CheckboxShape> blockShapes) {

            int nextCol = Integer.MAX_VALUE;

            for (CheckboxShape shape : blockShapes) {
                if (shape == currentShape) {
                    continue;
                }

                if (shape.row == currentShape.row && shape.col > currentShape.col && shape.col < nextCol) {
                    nextCol = shape.col;
                }
            }

            return nextCol;
        }

        private static String readNearestRightText(
                Sheet sheet,
                int rowIndex,
                int startCol,
                int endCol,
                DataFormatter formatter) {

            if (rowIndex < 0) {
                return "";
            }

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return "";
            }

            String sameText = getCellText(sheet, rowIndex, startCol, formatter);
            if (!sameText.isEmpty()) {
                return sameText;
            }

            int limit = endCol == Integer.MAX_VALUE
                    ? Math.max(row.getLastCellNum(), startCol + 4)
                    : endCol;

            for (int col = startCol + 1; col < limit; col++) {
                String text = getCellText(sheet, rowIndex, col, formatter);
                if (!text.isEmpty()) {
                    return text;
                }
            }

            return "";
        }

        private static String readBoundedText(
                Sheet sheet,
                int rowIndex,
                int startCol,
                int endCol,
                DataFormatter formatter) {

            if (rowIndex < 0) {
                return "";
            }

            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return "";
            }

            int limit = endCol == Integer.MAX_VALUE
                    ? Math.max(row.getLastCellNum(), startCol + 2)
                    : endCol;

            String sameText = getCellText(sheet, rowIndex, startCol, formatter);
            if (!sameText.isEmpty()) {
                return sameText;
            }

            for (int col = startCol + 1; col < limit; col++) {
                String text = getCellText(sheet, rowIndex, col, formatter);
                if (!text.isEmpty()) {
                    return text;
                }
            }

            return "";
        }

        private static void setChecked(Element shape, boolean checked) {

            Element clientData = firstChild(shape, EXCEL_NS, "ClientData");
            if (clientData == null) {
                return;
            }

            Element checkedNode = firstChild(clientData, EXCEL_NS, "Checked");

            if (checked) {
                if (checkedNode == null) {
                    checkedNode = shape.getOwnerDocument().createElementNS(EXCEL_NS, "x:Checked");
                    clientData.appendChild(checkedNode);
                }
                checkedNode.setTextContent("1");
            } else if (checkedNode != null) {
                clientData.removeChild(checkedNode);
            }
        }

        private static void setChecked(CtrlPropDoc ctrlPropDoc, boolean checked) {

            if (ctrlPropDoc == null || ctrlPropDoc.document == null) {
                return;
            }

            Element root = ctrlPropDoc.document.getDocumentElement();
            if (root == null || !CTRL_PROP_NS.equals(root.getNamespaceURI())) {
                return;
            }

            if (checked) {
                root.setAttribute("checked", "Checked");
            } else {
                root.removeAttribute("checked");
            }
        }

        private static Element firstChild(Element parent, String namespace, String localName) {

            NodeList nodes = parent.getChildNodes();

            for (int i = 0; i < nodes.getLength(); i++) {
                Node node = nodes.item(i);
                if (node instanceof Element) {
                    Element element = (Element) node;
                    if (localName.equals(element.getLocalName())
                            && namespace.equals(element.getNamespaceURI())) {
                        return element;
                    }
                }
            }

            return null;
        }
    }

    private static final class CheckboxShape {

        private final int col;
        private final int row;
        private final Element shapeElement;
        private final CtrlPropDoc ctrlPropDoc;

        private CheckboxShape(
                int col,
                int row,
                Element shapeElement,
                CtrlPropDoc ctrlPropDoc) {
            this.col = col;
            this.row = row;
            this.shapeElement = shapeElement;
            this.ctrlPropDoc = ctrlPropDoc;
        }
    }

    private static final class CtrlPropDoc {

        private final PackagePart part;
        private final Document document;

        private CtrlPropDoc(PackagePart part, Document document) {
            this.part = part;
            this.document = document;
        }
    }
}