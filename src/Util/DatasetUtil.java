package Util;

import java.io.File;

import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCSession;

public class DatasetUtil {

    private static final String RELATION_INFORMATION = "IMAN_specification";
    private static final String EXCEL_DATASET_TYPE = "MSExcelX";
    private static final String EXCEL_TOOL = "MSExcel";
    private static final String EXCEL_NAMED_REFERENCE = "excel";

    public static TCComponentDataset findDataset(
            TCComponentItemRevision itemRev,
            String datasetName) throws Exception {

        TCComponent[] related = itemRev.getRelatedComponents();

        for (TCComponent comp : related) {
            if (comp instanceof TCComponentDataset) {
                TCComponentDataset ds = (TCComponentDataset) comp;
                String name = ds.getProperty("object_name");

                if (name != null && name.trim().equalsIgnoreCase(datasetName.trim())) {
                    return ds;
                }
            }
        }

        return null;
    }

    public static File downloadDataset(TCComponentDataset dataset) throws Exception {

        TCComponentTcFile[] files = dataset.getTcFiles();

        if (files == null || files.length == 0) {
            throw new Exception("No files in dataset");
        }

        TCComponentTcFile tcFile = files[0];
        String fileName = tcFile.getProperty("original_file_name");

        File tempDir = new File("C:\\Temp");
        if (!tempDir.exists()) {
            tempDir.mkdirs();
        }

        tcFile.getFile(tempDir.getAbsolutePath());

        File downloaded = new File(tempDir, fileName);

        if (!downloaded.exists()) {
            throw new Exception("Dataset download failed");
        }

        return downloaded;
    }

    public static void uploadExcelToInformation(
            TCComponentItemRevision itemRev,
            File excelFile,
            String datasetName) throws Exception {

        if (excelFile == null || !excelFile.exists()) {
            throw new Exception("Exported excel file not found for upload");
        }

        TCComponentDataset dataset = findDatasetByRelationAndName(
                itemRev,
                RELATION_INFORMATION,
                datasetName);

        if (dataset == null) {
            dataset = createExcelDataset(itemRev, datasetName);
            itemRev.add(RELATION_INFORMATION, dataset);
        }

        dataset.setFiles(
                new String[] { excelFile.getAbsolutePath() },
                new String[] { EXCEL_NAMED_REFERENCE });
    }

    private static TCComponentDataset findDatasetByRelationAndName(
            TCComponentItemRevision itemRev,
            String relationName,
            String datasetName) throws Exception {

        TCComponent[] related = itemRev.getRelatedComponents(relationName);
        if (related == null) {
            return null;
        }

        for (TCComponent comp : related) {
            if (comp instanceof TCComponentDataset) {
                TCComponentDataset ds = (TCComponentDataset) comp;
                String name = ds.getProperty("object_name");

                if (name != null && name.trim().equalsIgnoreCase(datasetName.trim())) {
                    return ds;
                }
            }
        }

        return null;
    }

    private static TCComponentDataset createExcelDataset(
            TCComponentItemRevision itemRev,
            String datasetName) throws Exception {

        TCSession session = (TCSession) itemRev.getSession();

        TCComponentDatasetType datasetType =
                (TCComponentDatasetType) session.getTypeComponent("Dataset");

        return datasetType.create(
                datasetName,
                "",
                EXCEL_DATASET_TYPE,
                EXCEL_TOOL);
    }
}
