package mx.gob.imss.cit.mspmccifrascontrol.controller;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import mx.gob.imss.cit.mspmccommons.integration.model.CifrasControlMovimientosResponseDTO;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseStatus;
import org.springframework.web.bind.annotation.RestController;
import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccifrascontrol.services.MsPmcCifrasControlService;
import mx.gob.imss.cit.mspmccifrascontrol.services.ReporteService;
import mx.gob.imss.cit.mspmccommons.dto.ErrorResponse;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;
import net.sf.jasperreports.engine.JRException;

@RestController
@RequestMapping("/cifrascontrol/v1")
public class MsPmcCifrasControlController {

	private final Logger logger = LoggerFactory.getLogger(getClass());
	private String msg = "mspmccapados service ready to return";

	@Autowired
	private MsPmcCifrasControlService msPmcCifrasControlService;

	@Autowired
	ReporteService reportService;

	@RequestMapping("/health/ready")
	@ResponseStatus(HttpStatus.OK)
	public void ready() {
		//Indica que el ms esta listo para recibir peticiones
	}

	@RequestMapping("/health/live")
	@ResponseStatus(HttpStatus.OK)
	public void live() {
		//Indica que el ms esta vivo
	}

	@CrossOrigin(origins = "*", allowedHeaders = "*")
	@PostMapping("/cifrascontrol")
	public Object getCifrasOriginales(@RequestBody MsPmcCifrasControlInput input) {

		Object respuesta = null;

		logger.debug(msg);

		List<CifrasControlMovimientosResponseDTO> model;

		try {
			model = msPmcCifrasControlService.getCifrasControl(input);

			respuesta = new ResponseEntity<List<CifrasControlMovimientosResponseDTO>>(model, HttpStatus.OK);

		} catch (BusinessException be) {

			ErrorResponse errorResponse = be.getErrorResponse();

			int numberHTTPDesired = Integer.parseInt(errorResponse.getCode());

			respuesta = new ResponseEntity<ErrorResponse>(errorResponse, HttpStatus.valueOf(numberHTTPDesired));

		}

		return respuesta;
	}

	@CrossOrigin(origins = "*", allowedHeaders = "*")
	@PostMapping("/reportpdf")
	public Object getCifrasOriginalesReport(@RequestBody MsPmcCifrasControlInput input) throws BusinessException {

		Object respuesta = null;

		logger.debug(msg);

		Object model;

		try {
			model = reportService.getCifrasControlReport(input);

			respuesta = new ResponseEntity<Object>(model, HttpStatus.OK);

		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (JRException jre) {
			jre.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}

		return respuesta;
	}

	@CrossOrigin(origins = "*", allowedHeaders = "*")
	@PostMapping("/reportxls")
	public Object getCifrasOriginalesReportXls(@RequestBody MsPmcCifrasControlInput input) {

		Object respuesta = null;
		logger.debug(msg);
				
		try {

			HttpHeaders headers = new HttpHeaders();
			headers.add("Content-Disposition", "attachment; filename=CifrasControl.xlsx");
			headers.add("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
			Workbook wb = reportService.getCifrasControlReportXls(input);

			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			wb.write(bos);
			bos.close();
			respuesta = new ResponseEntity<Object>(bos.toByteArray(), HttpStatus.OK);
			return respuesta;

		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (JRException jre) {
			jre.printStackTrace();
		} catch(IOException ioe) {
			ioe.printStackTrace();
		} catch (BusinessException be) {
			ErrorResponse errorResponse = be.getErrorResponse();
			int numberHTTPDesired = Integer.parseInt(errorResponse.getCode());
			respuesta = new ResponseEntity<ErrorResponse>(errorResponse, HttpStatus.valueOf(numberHTTPDesired));
		}

		return respuesta;
	}

}
