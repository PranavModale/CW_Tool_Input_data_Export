package Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItemRevision;

import Util.DatasetUtil;
import Util.ExcelUtil;

public class ExportProgressiveToolToExcelService {

    private static final String TEMPLATE_DATASET_NAME = "Input Data Sheet for Progressive Tool";
    private static final String OUTPUT_DATASET_NAME = "Input Data Sheet for Progressive Tool.xlsx";
    private static final String ALLOWED_OBJECT_TYPE = "Tooling Assembly Revision";

    public static void execute() {

        FileInputStream fis = null;
        FileOutputStream fos = null;
        XSSFWorkbook workbook = null;

        try {
            InterfaceAIFComponent selected =
                    AIFUtility.getCurrentApplication().getTargetComponent();

            if (!(selected instanceof TCComponentItemRevision)) {
                showError(
                        "Export Progressive Tool Data to Excel",
                        "Please select a Tooling Assembly Revision that contains the dataset \""
                                + TEMPLATE_DATASET_NAME + "\".");
                return;
            }

            TCComponentItemRevision itemRev = (TCComponentItemRevision) selected;

            String objectType = itemRev.getProperty("object_type");
            if (objectType == null || !ALLOWED_OBJECT_TYPE.equalsIgnoreCase(objectType.trim())) {
                showError(
                        "Export Progressive Tool Data to Excel",
                        "Please select object type \"" + ALLOWED_OBJECT_TYPE + "\".");
                return;
            }

            TCComponentDataset dataset = DatasetUtil.findDataset(itemRev, TEMPLATE_DATASET_NAME);
            if (dataset == null) {
                showError(
                        "Export Progressive Tool Data to Excel",
                        "Selected object does not contain the dataset \""
                                + TEMPLATE_DATASET_NAME
                                + "\".\n\nPlease select the correct revision.");
                return;
            }

            File templateFile = DatasetUtil.downloadDataset(dataset);

            fis = new FileInputStream(templateFile);
            workbook = new XSSFWorkbook(fis);
            fis.close();
            fis = null;

            Sheet sheet = workbook.getSheetAt(0);

            itemRev.refresh();

            ExcelUtil.processSheet(sheet, itemRev, getMapping());
            ExcelUtil.processTableBlocks(sheet, itemRev, getTableMapping());

            File outDir = new File("C:\\Temp");
            if (!outDir.exists()) {
                outDir.mkdirs();
            }

            File tempOutput = new File(
                    outDir,
                    "ProgressiveToolData_temp_" + System.currentTimeMillis() + ".xlsx");
            File finalOutput = new File(outDir, OUTPUT_DATASET_NAME);

            fos = new FileOutputStream(tempOutput);
            workbook.write(fos);
            fos.close();
            fos = null;

            workbook.close();
            workbook = null;

            Files.move(
                    tempOutput.toPath(),
                    finalOutput.toPath(),
                    StandardCopyOption.REPLACE_EXISTING);

            if (!templateFile.getAbsolutePath().equalsIgnoreCase(finalOutput.getAbsolutePath())) {
                try {
                    Files.deleteIfExists(templateFile.toPath());
                } catch (Exception ex) {
                    System.out.println(
                            "Warning: could not delete old downloaded template: "
                                    + templateFile.getAbsolutePath());
                }
            }

            DatasetUtil.uploadExcelToInformation(itemRev, finalOutput, OUTPUT_DATASET_NAME);

            try {
                Files.deleteIfExists(finalOutput.toPath());
            } catch (Exception ex) {
                System.out.println(
                        "Warning: uploaded successfully, but could not delete local file: "
                                + finalOutput.getAbsolutePath());
            }

            MessageDialog.openInformation(
                    getShell(),
                    "Export Progressive Tool Data to Excel",
                    "Excel exported and uploaded successfully.");

        } catch (Exception e) {
            e.printStackTrace();
            showError(
                    "Export Progressive Tool Data to Excel",
                    "Export failed.\n\n" + e.getMessage());

        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (fos != null) {
                    fos.close();
                }
            } catch (Exception ignored) {
            }

            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (Exception ignored) {
            }
        }
    }

    private static void showError(String title, String message) {
        MessageDialog.openError(getShell(), title, message);
    }

    private static Shell getShell() {
        if (Display.getDefault() == null) {
            return null;
        }
        return Display.getDefault().getActiveShell();
    }

    public static Map<String, String> getTableMapping() {

        Map<String, String> map = new HashMap<String, String>();

        // Tool Base
        map.put("Tool Base|Top Plate|Material", "a2TopPlateMaterial");
        map.put("Tool Base|Top Plate|Specific Vendor", "a2TopPlateSpecificVendor");
        map.put("Tool Base|Top Plate|Specific Heat Treatment process", "a2TopPlateSHTProcess");
        map.put("Tool Base|Top Plate|Hardness", "a2TopPlateHardness");
        map.put("Tool Base|Top Plate|Recommended thickness(mm)", "a2TopPlateRecomendThickness");

        map.put("Tool Base|Punch Back Plate|Material", "a2PunchBackPlateMaterial");
        map.put("Tool Base|Punch Back Plate|Specific Vendor", "a2PunchBackPlateSpecVendor");
        map.put("Tool Base|Punch Back Plate|Specific Heat Treatment process", "a2PunchBackPlateSHTProcess");
        map.put("Tool Base|Punch Back Plate|Hardness", "a2PunchBackPlateHardness");
        map.put("Tool Base|Punch Back Plate|Recommended thickness(mm)", "a2PunchBackPlateRecommended");

        map.put("Tool Base|Punch Plate|Material", "a2PunchPlateMaterial");
        map.put("Tool Base|Punch Plate|Specific Vendor", "a2PunchPlateSpecificVendor");
        map.put("Tool Base|Punch Plate|Specific Heat Treatment process", "a2PunchPlateSHTProcess");
        map.put("Tool Base|Punch Plate|Hardness", "a2PunchPlateHardness");
        map.put("Tool Base|Punch Plate|Recommended thickness(mm)", "a2PunchPlateRecomthickness");

        map.put("Tool Base|Stripper Back Plate|Material", "a2StripperBackPlateMaterial");
        map.put("Tool Base|Stripper Back Plate|Specific Vendor", "a2StripperBckPlateSpeVendor");
        map.put("Tool Base|Stripper Back Plate|Specific Heat Treatment process", "a2StriperBckPlateSHTProcess");
        map.put("Tool Base|Stripper Back Plate|Hardness", "a2StripperBackPlateHardness");
        map.put("Tool Base|Stripper Back Plate|Recommended thickness(mm)", "a2StriperBkPlateRecThicknes");

        map.put("Tool Base|Stripper Plate|Material", "a2StripperPlateMaterial");
        map.put("Tool Base|Stripper Plate|Specific Vendor", "a2StripperPlateSpecVendor");
        map.put("Tool Base|Stripper Plate|Specific Heat Treatment process", "a2StripperPlateSHTProcess");
        map.put("Tool Base|Stripper Plate|Hardness", "a2StripperPlateHardness");
        map.put("Tool Base|Stripper Plate|Recommended thickness(mm)", "a2StriperPlateRecomthicknes");

        map.put("Tool Base|Die Plate|Material", "a2DiePlateMaterial");
        map.put("Tool Base|Die Plate|Specific Vendor", "a2DiePlateSpecificVendor");
        map.put("Tool Base|Die Plate|Specific Heat Treatment process", "a2DiePlateSHTProcess");
        map.put("Tool Base|Die Plate|Hardness", "a2DiePlateHardness");
        map.put("Tool Base|Die Plate|Recommended thickness(mm)", "a2DiePlateRecomThickness");

        map.put("Tool Base|Die Back Plate|Material", "a2DieBackPlateMaterial");
        map.put("Tool Base|Die Back Plate|Specific Vendor", "a2DieBackPlateSpecVendor");
        map.put("Tool Base|Die Back Plate|Specific Heat Treatment process", "a2DieBackPlateSHTProcess");
        map.put("Tool Base|Die Back Plate|Hardness", "a2DieBackPlateHardness");
        map.put("Tool Base|Die Back Plate|Recommended thickness(mm)", "a2DieBkPlateRecomThickness");

        map.put("Tool Base|Bottom Plate|Material", "a2BottomPlateMaterial");
        map.put("Tool Base|Bottom Plate|Specific Vendor", "a2BottomPlateSpecificVendor");
        map.put("Tool Base|Bottom Plate|Specific Heat Treatment process", "a2BottomPlateSHTProcess");
        map.put("Tool Base|Bottom Plate|Hardness", "a2BottomPlateHardness");
        map.put("Tool Base|Bottom Plate|Recommended thickness(mm)", "a2BottomPlateRecomThickness");

        map.put("Tool Base|Parallel Block|Material", "a2ParallelBlockMaterial");
        map.put("Tool Base|Parallel Block|Specific Vendor", "a2ParallelBlockSpecVendor");
        map.put("Tool Base|Parallel Block|Specific Heat Treatment process", "a2ParallelBlockSHTProcess");
        map.put("Tool Base|Parallel Block|Hardness", "a2ParallelBlockHardness");
        map.put("Tool Base|Parallel Block|Recommended thickness(mm)", "a2ParallelBlockRecoThicknes");

        // Punches & Die Insert
        map.put("Punches & Die Insert|Cutting Punch & Die Inserts|Material", "a2CuttingPDInsertMaterial");
        map.put("Punches & Die Insert|Cutting Punch & Die Inserts|Specific Vendor", "a2CuttingPDInsertSpecVendor");
        map.put("Punches & Die Insert|Cutting Punch & Die Inserts|Specific Heat Treatment process", "a2CuttingPDInsertSHTProcess");
        map.put("Punches & Die Insert|Cutting Punch & Die Inserts|Hardness/Remark", "a2CuttingPDIHardnessRemark");

        map.put("Punches & Die Insert|Bending Punch & Die Inserts|Material", "a2BendingPDInsertsMaterial");
        map.put("Punches & Die Insert|Bending Punch & Die Inserts|Specific Vendor", "a2BendingPDInsertSpecVendor");
        map.put("Punches & Die Insert|Bending Punch & Die Inserts|Specific Heat Treatment process", "a2BendingPDInsertSHTProcess");
        map.put("Punches & Die Insert|Bending Punch & Die Inserts|Hardness/Remark", "a2BendingPDIHardnessRemark");

        return map;
    }

    public static Map<String, String> getMapping() {

        Map<String, String> map = new HashMap<String, String>();

        // Input Data Sheet For Progressive Tool
        map.put("Project No", "a2ProjectNo");
        map.put("Project Name", "a2CatalogueNoComp_P");
        map.put("Sub- Project No.", "a2SubProjectNo");
        map.put("Sub-Project Name", "a2SubProjectName");
        map.put("Manufacturer", "a2Manufacturer");
        map.put("Project Type", "a2ProjectType");
        map.put("Date", "a2Date");
        map.put("Prepared By", "a2PreparedBy");

        // PART
        map.put("Part Name", "a2PartNameSMP");
        map.put("PLM ID/Drawing No. & Revision", "MULTI_PLM_DRAWING_REV");
        map.put("Part Weight", "a2PartWeightSMP");
        map.put("Raw Material", "a2RawMaterialSMP");
        map.put("Coil Size", "a2CoilSize");
        map.put("Pitch", "a2Pitch");
        map.put("Burr Direction", "a2BurrDirection");
        map.put("Version Change-Over(If any)", "a2VersionChangeOver_Part");
        map.put("Additional Note (If any)", "a2AdditionalNote_Part");

        // TOOL
        map.put("Tool Type", "a2Tooltype");
        map.put("No of impression", "a2Noofimpressions");
        map.put("Expected tool Life", "a2ExpectedToolLife_New");
        map.put("Press Machine", "a2PressMachine");

        // TOOL BASE
        map.put("Std. Boughtout items for Tool Base( eg. guiding elements,sensors, etc.)", "a2StdBoughtoutItemsToolBase");

        // Ejection
        map.put("Part collected on", "a2PartCollectedOn_New");

        // TOOL SECURITY
        map.put("Tool Security", "a2ToolSecurity_New");

        // Design (CAD Format)
        map.put("For In-house", "a2CADDesignFormatForInhouse");
        map.put("For Outsource", "a2CADDesignFormatOutsource");

        return map;
    }
}