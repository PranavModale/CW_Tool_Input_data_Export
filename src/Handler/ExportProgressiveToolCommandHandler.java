package Handler;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;

import Service.ExportProgressiveToolToExcelService;

public class ExportProgressiveToolCommandHandler extends AbstractHandler {

    @Override
    public Object execute(ExecutionEvent event) {
        ExportProgressiveToolToExcelService.execute();
        return null;
    }
}
