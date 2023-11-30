package mx.gob.imss.cit.mspmccifrascontrol.services;

import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;
import net.sf.jasperreports.engine.JRException;

public interface ReporteService {

	Object getCifrasControlReport(MsPmcCifrasControlInput input)
			throws JRException, IOException, BusinessException;

	Workbook getCifrasControlReportXls(MsPmcCifrasControlInput input)
			throws JRException, IOException, BusinessException;

}
