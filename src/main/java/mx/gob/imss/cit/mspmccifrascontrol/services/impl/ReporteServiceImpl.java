package mx.gob.imss.cit.mspmccifrascontrol.services.impl;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.Month;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.Base64;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;

import mx.gob.imss.cit.mspmccommons.integration.model.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccifrascontrol.integration.dao.ParametroRepository;
import mx.gob.imss.cit.mspmccifrascontrol.services.MsPmcCifrasControlService;
import mx.gob.imss.cit.mspmccifrascontrol.services.ReporteService;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;
import mx.gob.imss.cit.mspmccommons.utils.DateUtils;
import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JRException;
import net.sf.jasperreports.engine.JasperCompileManager;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.JasperReport;

@Service("reportService")
public class ReporteServiceImpl implements ReporteService {

	private static final String NOMBRE_INST = "nombreInstitucion";
	private static final String DIRECCION_INST = "direccionInstitucion";
	private static final String UNIDAD_INST = "unidadInstitucion";
	private static final String DIVISION_INST = "divisionInstitucion";
	private static final String NOMBRE_REPORTE = "nombreReporte";
	private static final String REPORT = "Report Generatade succefully";
	private static final String REPORTE_DEL = "reporteDelegacional";
	private static final String REPORTE_NAL = "reporteNacional";
	private static final String REGISTROS = "Registros";
	private static final String CORRECTOS = "Correctos";
	private static final String ERRONEOS = "Erróneos";
	private static final String DUPLICADOS = "Duplicados";
	private static final String SUS_AJUSTE = "Susceptibles de Ajuste";
	
	public String exportReport() { 
		return REPORT;
	}

	@Autowired
	private MsPmcCifrasControlService cifrasControlService;

	@Autowired
	private ParametroRepository parametroRepository;

	DateTimeFormatter europeanDateFormatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

	@Override
	public Object getCifrasControlReport(MsPmcCifrasControlInput input)
			throws JRException, IOException, BusinessException {

		List<CifrasControlMovimientosResponseDTO> cifrasControl = cifrasControlService.getCifrasControl(input);
		CifrasControlDTO cifrasControlDTO = getCifrasTotales(cifrasControl);

		Optional<ParametroDTO> nombreInstitucion = parametroRepository.findOneByCve(NOMBRE_INST);
		Optional<ParametroDTO> direccionInstitucion = parametroRepository.findOneByCve(DIRECCION_INST);
		Optional<ParametroDTO> unidadInstitucion = parametroRepository.findOneByCve(UNIDAD_INST);
		Optional<ParametroDTO> coordinacionInstituc = parametroRepository.findOneByCve("coordinacionInstitucion");
		Optional<ParametroDTO> divisionInstitucion = parametroRepository.findOneByCve(DIVISION_INST);
		Optional<ParametroDTO> nombreReporte = parametroRepository.findOneByCve(NOMBRE_REPORTE);

		Map<String, Object> parameters = new HashMap<>();

		InputStream resourceAsStream = null;
		if (Boolean.TRUE.equals(input.getDelRegPat())){
			resourceAsStream = ReporteServiceImpl.class.getResourceAsStream("/cifrasControlDelegacion.jrxml");
			Optional<ParametroDTO> reporteDelegacional = parametroRepository.findOneByCve(REPORTE_DEL);
			String repDeleg = "";
			if(reporteDelegacional.isPresent()) {
				repDeleg = reporteDelegacional.get().getDesParametro();
			}
			parameters.put(REPORTE_DEL, repDeleg);
		} else {
			resourceAsStream = ReporteServiceImpl.class.getResourceAsStream("/cifrasControlNacional.jrxml");
			Optional<ParametroDTO> reporteNacional = parametroRepository.findOneByCve(REPORTE_NAL);
			String repNacional = "";
			if(reporteNacional.isPresent()) {
				repNacional = reporteNacional.get().getDesParametro();
			}
			parameters.put(REPORTE_DEL, repNacional);
		}

		JasperReport jasperReport = JasperCompileManager.compileReport(resourceAsStream);
		
		String institucion = "";
		String direccion = "";
		String unidad = "";
		String coordinacion = "";
		String division = "";
		String nombre = "";
		if(nombreInstitucion.isPresent() && direccionInstitucion.isPresent() && unidadInstitucion.isPresent() && 
		   coordinacionInstituc.isPresent() && divisionInstitucion.isPresent() && nombreReporte.isPresent()) {
			institucion = nombreInstitucion.get().getDesParametro();
			direccion = direccionInstitucion.get().getDesParametro();
			unidad = unidadInstitucion.get().getDesParametro();
			coordinacion = coordinacionInstituc.get().getDesParametro();
			division = divisionInstitucion.get().getDesParametro();
			nombre = nombreReporte.get().getDesParametro();
		}
		parameters.put(NOMBRE_INST, institucion);
		parameters.put(DIRECCION_INST, direccion);
		parameters.put(UNIDAD_INST, unidad);
		parameters.put("coordinacionInstituc", coordinacion);
		parameters.put(DIVISION_INST, division);
		parameters.put(NOMBRE_REPORTE, nombre);
		parameters.put("numTotalRegistros", cifrasControlDTO.getNumTotalRegistros());
		parameters.put("numRegistrosCorrectos", cifrasControlDTO.getNumRegistrosCorrectos());
		parameters.put("numRegistrosCorrectosOtras", cifrasControlDTO.getNumRegistrosCorrectosOtras());
		parameters.put("numRegistrosErrorOtras", cifrasControlDTO.getNumRegistrosErrorOtras());
		parameters.put("numRegistrosError", cifrasControlDTO.getNumRegistrosError());
		parameters.put("numRegistrosSusOtras", cifrasControlDTO.getNumRegistrosSusOtras());
		parameters.put("numRegistrosDupOtras", cifrasControlDTO.getNumRegistrosDupOtras());
		parameters.put("numRegistrosDup", cifrasControlDTO.getNumRegistrosDup());
		parameters.put("numRegistrosSus", cifrasControlDTO.getNumRegistrosSus());

		parameters.put("fromDate",
				DateUtils.calcularFechaPoceso(input.getFromMonth(), input.getFromYear()).format(europeanDateFormatter));
		parameters.put("toDate",
				DateUtils.calcularFechaPocesoFin(input.getToMonth(), input.getToYear()).format(europeanDateFormatter));

		parameters.put("cifrasDataSource", cifrasControl);

		JasperPrint print = JasperFillManager.fillReport(jasperReport, parameters, new JREmptyDataSource());
		return Base64.getEncoder().encodeToString(JasperExportManager.exportReportToPdf(print));
	}

	private CifrasControlDTO getCifrasTotales(List<CifrasControlMovimientosResponseDTO> cifrasControl) {
		return cifrasControl.stream().map(ccm -> {
			CifrasControlDTO cc = new CifrasControlDTO();
			cc.setNumRegistrosCorrectos(ccm.getCorrecto());
			cc.setNumRegistrosError(ccm.getErroneo());
			cc.setNumRegistrosDup(ccm.getDuplicado());
			cc.setNumRegistrosSus(ccm.getSusceptible());
			cc.setNumRegistrosCorrectosOtras(ccm.getCorrectoOtras());
			cc.setNumRegistrosErrorOtras(ccm.getErroneoOtras());
			cc.setNumRegistrosDupOtras(ccm.getDuplicadoOtras());
			cc.setNumRegistrosSusOtras(ccm.getSusceptibleOtras());
			cc.setNumTotalRegistros(ccm.getTotal());
			return cc;
		}).reduce(new CifrasControlDTO(), (prev, curr) -> {
			CifrasControlDTO cifrasControlDTO = new CifrasControlDTO();
			cifrasControlDTO.setNumRegistrosCorrectos(prev.getNumRegistrosCorrectos() + curr.getNumRegistrosCorrectos());
			cifrasControlDTO.setNumRegistrosError(prev.getNumRegistrosError() + curr.getNumRegistrosError());
			cifrasControlDTO.setNumRegistrosDup(prev.getNumRegistrosDup() + curr.getNumRegistrosDup());
			cifrasControlDTO.setNumRegistrosSus(prev.getNumRegistrosSus() + curr.getNumRegistrosSus());
			cifrasControlDTO.setNumRegistrosCorrectosOtras(prev.getNumRegistrosCorrectosOtras() + curr.getNumRegistrosCorrectosOtras());
			cifrasControlDTO.setNumRegistrosErrorOtras(prev.getNumRegistrosErrorOtras() + curr.getNumRegistrosErrorOtras());
			cifrasControlDTO.setNumRegistrosDupOtras(prev.getNumRegistrosDupOtras() + curr.getNumRegistrosDupOtras());
			cifrasControlDTO.setNumRegistrosSusOtras(prev.getNumRegistrosSusOtras() + curr.getNumRegistrosSusOtras());
			cifrasControlDTO.setNumTotalRegistros(prev.getNumTotalRegistros() + curr.getNumTotalRegistros());
			return cifrasControlDTO;
		});
	}
	
	private void region(CellRangeAddress region, Sheet sheetDelegacional, CellStyle headerStyle) {
		for (int x = region.getFirstRow(); x < region.getLastRow(); x++) {
			Row row = sheetDelegacional.createRow(x);
			for (int y = region.getFirstColumn(); y < region.getLastColumn(); y++) {
				Cell cell = row.createCell(y);
				cell.setCellValue(" ");
				cell.setCellStyle(headerStyle);
			}
		}
	}
	
	private void regionNacional(MsPmcCifrasControlInput input, Sheet sheetNacional, CellStyle headerStyle) {
		CellRangeAddress region = Boolean.TRUE.equals(input.getDelRegPat()) ? CellRangeAddress.valueOf("B2:M7")
				: CellRangeAddress.valueOf("B2:L7");
		for (int i = region.getFirstRow(); i < region.getLastRow(); i++) {
			Row row = sheetNacional.createRow(i);
			for (int j = region.getFirstColumn(); j < region.getLastColumn(); j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue(" ");
				cell.setCellStyle(headerStyle);
			}
		}
	}
	
	private String nombreInstitucion(Optional<ParametroDTO> nombreInstitucion) {
		
		return nombreInstitucion.isPresent() ? nombreInstitucion.get().getDesParametro() : "";
		
	}
	
	private String direccionInstitucion(Optional<ParametroDTO> direccionInstitucion) {
		
		return direccionInstitucion.isPresent() ? direccionInstitucion.get().getDesParametro() : "";

	}
	
	private String unidadInstitucion(Optional<ParametroDTO> unidadInstitucion) {
		
		return unidadInstitucion.isPresent() ? unidadInstitucion.get().getDesParametro() : "";
		
	}
	
	private String coordinacionInstitucion(Optional<ParametroDTO> coordinacionInstituc) {
		
		return coordinacionInstituc.isPresent() ? coordinacionInstituc.get().getDesParametro() : "";
		
	}
	
	private String divisionInstitucion(Optional<ParametroDTO> divisionInstitucion) {
		
		return divisionInstitucion.isPresent() ? divisionInstitucion.get().getDesParametro() : "";
		
	}

	@Override
	public Workbook getCifrasControlReportXls(MsPmcCifrasControlInput input)
			throws JRException, IOException, BusinessException {
		List<CifrasControlMovimientosResponseDTO> cifrasControl = cifrasControlService.getCifrasControl(input);
		CifrasControlDTO cifrasControlDTO = getCifrasTotales(cifrasControl);
		Optional<ParametroDTO> nombreInstitucion = parametroRepository.findOneByCve(NOMBRE_INST);
		Optional<ParametroDTO> direccionInstitucion = parametroRepository.findOneByCve(DIRECCION_INST);
		Optional<ParametroDTO> unidadInstitucion = parametroRepository.findOneByCve(UNIDAD_INST);
		Optional<ParametroDTO> coordinacionInstituc = parametroRepository.findOneByCve("coordinacionInstitucion");
		Optional<ParametroDTO> divisionInstitucion = parametroRepository.findOneByCve(DIVISION_INST);
		Optional<ParametroDTO> nombreReporte = parametroRepository.findOneByCve(NOMBRE_REPORTE);
		Workbook workbook = new XSSFWorkbook();
		XSSFFont font = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Montserrat");
		font.setFontHeightInPoints((short) 8);
		font.setBold(true);

		XSSFFont fontPeriodo = ((XSSFWorkbook) workbook).createFont();
		fontPeriodo.setFontName("Montserrat");
		fontPeriodo.setFontHeightInPoints((short) 8);
		fontPeriodo.setColor(HSSFColor.WHITE.index);
		fontPeriodo.setBold(true);
		CellStyle headerStyle = createStyle(font, HorizontalAlignment.LEFT, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.BLACK.index, workbook, false);
		LocalDate localDate = LocalDate.now(ZoneId.of("America/Mexico_City"));
		DateTimeFormatter df = DateTimeFormatter.ofPattern("dd/MM/yyyy", Locale.getDefault());
		Locale mexico = new Locale("es", "MX");
		InputStream inputStream = ReporteServiceImpl.class.getResourceAsStream("/IMSS-logo-.png");
		byte[] bytes = IOUtils.toByteArray(inputStream);
		int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		inputStream.close();
		// Returns an object that handles instantiating concrete classes
		CreationHelper helper = workbook.getCreationHelper();

		if (Boolean.TRUE.equals(input.getDelRegPat())) {
			List<Month> months = processMonths(input);
			for (int i = 0; i < months.size(); i++) {
				Month fromMonth = months.get(i);
				input.setFromMonth(String.valueOf(fromMonth.getValue()));
				input.setToMonth(String.valueOf(fromMonth.getValue()));

				if (!cifrasControl.isEmpty()) {
					Sheet sheetDelegacional = workbook.createSheet(
							"Delegacional " + fromMonth.getDisplayName(TextStyle.SHORT, mexico).toUpperCase());
					Header header = sheetDelegacional.getHeader();
					header.setRight("Hoja " + HeaderFooter.page() + " de " + HeaderFooter.numPages());

					sheetDelegacional.setColumnWidth(0, 200);
					sheetDelegacional.setMargin(Sheet.LeftMargin, 0.5 /* inches */ );
					sheetDelegacional.setMargin(Sheet.RightMargin, 0.5 /* inches */ );

					Drawing drawing = sheetDelegacional.createDrawingPatriarch();

					ClientAnchor anchor = helper.createClientAnchor();
					// set top-left corner for the image
					anchor.setCol1(1);
					anchor.setRow1(1);

					// Creates a picture
					Picture pict = drawing.createPicture(anchor, pictureIdx);
					// Reset the image to the original size
					pict.resize(1.5, 5);

					CellRangeAddress region = CellRangeAddress.valueOf("B2:M7");
					
					region(region, sheetDelegacional, headerStyle);

					Row rowInstitucionDet = sheetDelegacional.getRow(1);

					Cell cellInstitucionDet = rowInstitucionDet.getCell(3);
					
					String nomInst = nombreInstitucion(nombreInstitucion);
					
					cellInstitucionDet.setCellValue(nomInst);
					cellInstitucionDet.setCellStyle(headerStyle);

					Row rowDireccionDet = sheetDelegacional.getRow(2);
					Cell cellDireccionDet = rowDireccionDet.getCell(3);
					
					String dirInst = direccionInstitucion(direccionInstitucion);
					
					cellDireccionDet.setCellValue(dirInst);
					cellDireccionDet.setCellStyle(headerStyle);

					Row rowUnidadDet = sheetDelegacional.getRow(3);
					Cell cellUnidadDet = rowUnidadDet.getCell(3);
					
					String uniInst = unidadInstitucion(unidadInstitucion);
					
					cellUnidadDet.setCellValue(uniInst);
					cellUnidadDet.setCellStyle(headerStyle);

					Row rowCoordinacionDet = sheetDelegacional.getRow(4);
					Cell cellCoordinacionDet = rowCoordinacionDet.getCell(3);
					
					String coorInst = coordinacionInstitucion(coordinacionInstituc);					
					
					cellCoordinacionDet.setCellValue(coorInst);
					cellCoordinacionDet.setCellStyle(headerStyle);
					
					String divInst = divisionInstitucion(divisionInstitucion);
					
					Row rowDivisionDet = sheetDelegacional.getRow(5);
					Cell cellDivisionDet = rowDivisionDet.getCell(3);
					cellDivisionDet.setCellValue(divInst);
					cellDivisionDet.setCellStyle(headerStyle);

					Cell cellFechaDet = rowDivisionDet.getCell(11);
					cellFechaDet.setCellValue(localDate.format(df));
					cellFechaDet.setCellStyle(headerStyle);

					sheetDelegacional.addMergedRegion(CellRangeAddress.valueOf("D2:I2"));
					sheetDelegacional.addMergedRegion(CellRangeAddress.valueOf("D3:J3"));
					sheetDelegacional.addMergedRegion(CellRangeAddress.valueOf("D4:J4"));
					sheetDelegacional.addMergedRegion(CellRangeAddress.valueOf("D5:J5"));
					sheetDelegacional.addMergedRegion(CellRangeAddress.valueOf("D6:J6"));

					if(Boolean.TRUE.equals(input.getDelRegPat())) {
						Optional<ParametroDTO> reporteDelegacional = parametroRepository.findOneByCve(REPORTE_DEL);
						createHeaderReport(input, nombreReporte, reporteDelegacional, sheetDelegacional, true, font,
								fontPeriodo, workbook, fromMonth.getDisplayName(TextStyle.FULL, mexico).toUpperCase());
					}else {
						Optional<ParametroDTO> reporteNacional = parametroRepository.findOneByCve(REPORTE_NAL);
						createHeaderReport(input, nombreReporte, reporteNacional, sheetDelegacional, true, font,
								fontPeriodo, workbook, fromMonth.getDisplayName(TextStyle.FULL, mexico).toUpperCase());
					}					

					int counterdet = createHeaderTable(cifrasControl, sheetDelegacional, font, workbook);

					fillDetailReport(cifrasControlDTO, sheetDelegacional, counterdet, true, font, workbook);
				}
			}

		} else {

			String sheetName = "Nacional";

			Sheet sheetNacional = workbook.createSheet(sheetName);
			sheetNacional.setColumnWidth(0, 200);

			Header header = sheetNacional.getHeader();
			header.setRight("Página " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

			sheetNacional.setMargin(Sheet.LeftMargin, 0.5 /* inches */ );

			sheetNacional.setMargin(Sheet.RightMargin, 0.5 /* inches */ );

			// Creates the top-level drawing patriarch.
			Drawing drawing = sheetNacional.createDrawingPatriarch();

			ClientAnchor anchor = helper.createClientAnchor();
			// set top-left corner for the image
			anchor.setCol1(1);
			anchor.setRow1(1);

			// Creates a picture
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			// Reset the image to the original size
			pict.resize(1.5, 5);
			
			regionNacional(input, sheetNacional, headerStyle);

			Row rowInstitucion = sheetNacional.getRow(1);

			Cell cellInstitucion = rowInstitucion.getCell(3);
			cellInstitucion.setCellValue(nombreInstitucion.get().getDesParametro());
			cellInstitucion.setCellStyle(headerStyle);

			Row rowDireccion = sheetNacional.getRow(2);
			Cell cellDireccion = rowDireccion.getCell(3);
			cellDireccion.setCellValue(direccionInstitucion.get().getDesParametro());
			cellDireccion.setCellStyle(headerStyle);

			Row rowUnidad = sheetNacional.getRow(3);
			Cell cellUnidad = rowUnidad.getCell(3);
			cellUnidad.setCellValue(unidadInstitucion.get().getDesParametro());
			cellUnidad.setCellStyle(headerStyle);

			Row rowCoordinacion = sheetNacional.getRow(4);
			Cell cellCoordinacion = rowCoordinacion.getCell(3);
			cellCoordinacion.setCellValue(coordinacionInstituc.get().getDesParametro());
			cellCoordinacion.setCellStyle(headerStyle);

			Row rowDivision = sheetNacional.getRow(5);
			Cell cellDivision = rowDivision.getCell(3);
			cellDivision.setCellValue(divisionInstitucion.get().getDesParametro());
			cellDivision.setCellStyle(headerStyle);

			Cell cellFecha = rowDivision.getCell(10);
			cellFecha.setCellValue(localDate.format(df));
			cellFecha.setCellStyle(headerStyle);

			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D2:I2"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D3:J3"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D4:J4"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D5:J5"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D6:J6"));

			if(Boolean.TRUE.equals(input.getDelRegPat())) {
				Optional<ParametroDTO> reporteDelegacional = parametroRepository.findOneByCve(REPORTE_DEL);
				createHeaderReport(input, nombreReporte, reporteDelegacional, sheetNacional, false, font, fontPeriodo, workbook,
						null);
			}else {
				Optional<ParametroDTO> reporteNacional = parametroRepository.findOneByCve(REPORTE_NAL);
				createHeaderReport(input, nombreReporte, reporteNacional, sheetNacional, false, font, fontPeriodo, workbook,
						null);
			}			

			int counter = createHeaderTable(cifrasControl, sheetNacional, font, workbook);

			fillDetailReport(cifrasControlDTO, sheetNacional, counter, false, font, workbook);
		}

		return workbook;
	}

	private List<Month> processMonths(MsPmcCifrasControlInput input) {
		LocalDate fecProcesoIni = DateUtils.calcularFecPoceso(input.getFromMonth(), input.getFromYear());
		LocalDate fecProcesoFin = DateUtils.calcularFecPocesoFin(input.getToMonth(), input.getToYear());
		List<Month> months = new ArrayList<>();
		int initialMonth = fecProcesoIni.getMonthValue();
		int finalMonth = fecProcesoFin.getMonthValue();
		for (int i = initialMonth; i <= finalMonth; i++) {
			months.add(Month.of(i));
		}

		return months;
	}

	private void fillDetailReport(CifrasControlDTO cifrasControlDTO, Sheet sheetNacional,
								  int counter, boolean delRegaPat, XSSFFont font, Workbook workbook) {
		Row rowTotal = sheetNacional.createRow(counter + 14);

		CellStyle rowColorStyle = createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.WHITE.index, false, HSSFColor.WHITE.index, workbook, true);
		CellStyle rowtyle = createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true);

		CellStyle style = null;
		if ((counter + 14) % 2 == 0) {
			style = rowtyle;
		} else {
			style = rowColorStyle;
		}

		Cell tipoArchivoDetCell = rowTotal.createCell(1);
		tipoArchivoDetCell.setCellValue("Totales");
		tipoArchivoDetCell.setCellStyle(style);

		if (delRegaPat) {

			Cell delegacionDetCell = rowTotal.createCell(2);
			delegacionDetCell.setCellValue("");
			delegacionDetCell.setCellStyle(style);

			fDetailReportDel(rowTotal, cifrasControlDTO, style);
			
		} else {
			
			fDetailReport(rowTotal, cifrasControlDTO, style);
				
		}
	}
	
	private CellStyle estilos(int i, CellStyle rowColorStyle, CellStyle rowtyle) {
		
		CellStyle style = null;
		
		if (i % 2 == 0) {
			style = rowColorStyle;
		} else {
			style = rowtyle;
		}
		
		return style;
	}
	
	private int del(CifrasControlMovimientosResponseDTO detalleConsultaDTO, Row rowDetail, CellStyle style) {
		int del = 0;
		if (detalleConsultaDTO.get_id() != null) {
			Cell delegacionCell = rowDetail.createCell(2);
			delegacionCell.setCellValue(detalleConsultaDTO.get_id());
			delegacionCell.setCellStyle(style);
			del = 1;
		}
		return del;
	}
	
	private void celdas(Row rowDetail, int del, CifrasControlMovimientosResponseDTO detalleConsultaDTO, CellStyle style) {
		Cell correctosDetCell = rowDetail.createCell(3 + del);
		correctosDetCell.setCellValue(detalleConsultaDTO.getCorrecto() != null
				? detalleConsultaDTO.getCorrecto()
				: 0);
		correctosDetCell.setCellStyle(style);

		Cell erroneosDetCell = rowDetail.createCell(4 + del);
		erroneosDetCell.setCellValue(
				detalleConsultaDTO.getErroneo() != null ? detalleConsultaDTO.getErroneo()
						: 0);
		erroneosDetCell.setCellStyle(style);

		Cell duplicadosDetCell = rowDetail.createCell(5 + del);
		duplicadosDetCell.setCellValue(
				detalleConsultaDTO.getDuplicado() != null ? detalleConsultaDTO.getDuplicado() : 0);
		duplicadosDetCell.setCellStyle(style);

		Cell susAjusDetCell = rowDetail.createCell(6 + del);
		susAjusDetCell.setCellValue(
				detalleConsultaDTO.getSusceptible() != null ? detalleConsultaDTO.getSusceptible() : 0);
		susAjusDetCell.setCellStyle(style);

		Cell correctosOtrasDetCell = rowDetail.createCell(7 + del);
		correctosOtrasDetCell.setCellValue(detalleConsultaDTO.getCorrectoOtras() != null
				? detalleConsultaDTO.getCorrectoOtras()
				: 0);
		correctosOtrasDetCell.setCellStyle(style);

		Cell erroneosOtrasDetCell = rowDetail.createCell(8 + del);
		erroneosOtrasDetCell.setCellValue(detalleConsultaDTO.getErroneoOtras() != null
				? detalleConsultaDTO.getErroneoOtras()
				: 0);
		erroneosOtrasDetCell.setCellStyle(style);

		Cell duplicadosOtrasDetCell = rowDetail.createCell(9 + del);
		duplicadosOtrasDetCell.setCellValue(detalleConsultaDTO.getDuplicadoOtras() != null
				? detalleConsultaDTO.getDuplicadoOtras()
				: 0);
		duplicadosOtrasDetCell.setCellStyle(style);

		Cell susAjusOtrasDetCell = rowDetail.createCell(10 + del);
		susAjusOtrasDetCell.setCellValue(detalleConsultaDTO.getSusceptibleOtras() != null
				? detalleConsultaDTO.getSusceptibleOtras()
				: 0);
		susAjusOtrasDetCell.setCellStyle(style);
	}

	private int createHeaderTable(List<CifrasControlMovimientosResponseDTO> detalleConsultaDTOList, Sheet sheetNacional,
			                      XSSFFont font, Workbook workbook) {
		int counter = 1;
		for (int i = 0; i < detalleConsultaDTOList.size(); i++) {
			CifrasControlMovimientosResponseDTO detalleConsultaDTO = detalleConsultaDTOList.get(i);
			counter++;
			Row rowDetail = sheetNacional.createRow(i + 15);

			CellStyle rowColorStyle = createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.WHITE.index, false, HSSFColor.WHITE.index, workbook, true);
			CellStyle rowtyle = createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true);

			CellStyle style = estilos(i, rowColorStyle, rowtyle);

			Cell tipoArchivoDetCell = rowDetail.createCell(1);
			tipoArchivoDetCell.setCellValue(detalleConsultaDTO.getCveOrigenArchivo());
			tipoArchivoDetCell.setCellStyle(style);

			int del = del(detalleConsultaDTO, rowDetail, style);

			Cell totalDetCell = rowDetail.createCell(2 + del);
			totalDetCell.setCellValue(detalleConsultaDTO.getTotal());
			totalDetCell.setCellStyle(style);

			celdas(rowDetail, del, detalleConsultaDTO, style);
			
		}
		return counter;
	}
	
	private String nombreReporte(Optional<ParametroDTO> nombreReporte) {
		
		return nombreReporte.isPresent() ? nombreReporte.get().getDesParametro() : "";
		
	}
	
	private String reporteNacional(Optional<ParametroDTO> reporteNacional) {
		
		return reporteNacional.isPresent() ? reporteNacional.get().getDesParametro() : "";

	}
	
	private void hojaNacional(MsPmcCifrasControlInput input, Sheet sheetNacional) {
		if (Boolean.TRUE.equals(input.getDelRegPat())) {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B8:L8"));
		} else {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B8:K8"));
		}
	}
	
	private void sheetNal(MsPmcCifrasControlInput input, Sheet sheetNacional) {
		if (Boolean.TRUE.equals(input.getDelRegPat())) {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B9:L9"));
		} else {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B9:K9"));
		}
	}

	private void createHeaderReport(MsPmcCifrasControlInput input, Optional<ParametroDTO> nombreReporte,
			Optional<ParametroDTO> reporteNacional, Sheet sheetNacional, boolean b, XSSFFont font, XSSFFont fontPeriodo,
			Workbook workbook, String monthName) {
		Row rowNombreReporte = sheetNacional.createRow(7);
		Cell nombreReporteCell = rowNombreReporte.createCell(1);
		
		String nomReporte = nombreReporte(nombreReporte);
		
		nombreReporteCell.setCellValue(nomReporte);
		nombreReporteCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.WHITE.index, false, HSSFColor.WHITE.index, workbook, true));
		
		hojaNacional(input, sheetNacional);

		Row rowReporteNacional = sheetNacional.createRow(8);
		Cell reporteNacionalCell = rowReporteNacional.createCell(1);
		
		String reporteNal = reporteNacional(reporteNacional);
		
		reporteNacionalCell.setCellValue(reporteNal);
		reporteNacionalCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.WHITE.index, false, HSSFColor.WHITE.index, workbook, true));
		
		sheetNal(input, sheetNacional);

		Row rowPeriodoConsultado = sheetNacional.createRow(10);
		Cell periodoConsultado = rowPeriodoConsultado.createCell(1);
		periodoConsultado.setCellValue("Periodo consultado: ");
		periodoConsultado.setCellStyle(createStyle(fontPeriodo, HorizontalAlignment.LEFT, VerticalAlignment.CENTER,
				HSSFColor.GREY_50_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B11:C11"));

		Cell fecInicio = rowPeriodoConsultado.createCell(3);
		fecInicio.setCellValue(
				DateUtils.calcularFechaPoceso(input.getFromMonth(), input.getFromYear()).format(europeanDateFormatter));
		fecInicio.setCellStyle(createStyle(fontPeriodo, HorizontalAlignment.LEFT, VerticalAlignment.CENTER,
				HSSFColor.GREY_50_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell al = rowPeriodoConsultado.createCell(4);
		al.setCellValue(" al ");
		al.setCellStyle(createStyle(fontPeriodo, HorizontalAlignment.LEFT, VerticalAlignment.CENTER,
				HSSFColor.GREY_50_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell fecFin = rowPeriodoConsultado.createCell(5);
		fecFin.setCellValue(
				DateUtils.calcularFechaPocesoFin(input.getToMonth(), input.getToYear()).format(europeanDateFormatter));
		fecFin.setCellStyle(createStyle(fontPeriodo, HorizontalAlignment.LEFT, VerticalAlignment.CENTER,
				HSSFColor.GREY_50_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Row rowRegistros = sheetNacional.createRow(12);
		Cell registrosCell = rowRegistros.createCell(3);
		registrosCell.setCellValue(REGISTROS);
		registrosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		sheetNacional.addMergedRegion(CellRangeAddress.valueOf("B13:C14"));

		Cell mesConsultado = rowRegistros.createCell(1);
		mesConsultado.setCellValue(monthName);
		mesConsultado.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		if (Boolean.TRUE.equals(input.getDelRegPat()) && b) {
			Cell registrosCellOtras = rowRegistros.createCell(8);
			registrosCellOtras.setCellValue(REGISTROS);
			registrosCellOtras.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		} else {
			Cell registrosCellOtras = rowRegistros.createCell(7);
			registrosCellOtras.setCellValue(REGISTROS);
			registrosCellOtras.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		}

		// Nacional
		if (Boolean.TRUE.equals(input.getDelRegPat()) && b) {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D13:H13"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("I13:L13"));
		} else {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D13:G13"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("H13:K13"));
		}

		Row rowDelegacion = sheetNacional.createRow(13);
		Cell delegacionCell = rowDelegacion.createCell(3);
		delegacionCell.setCellValue("Delegación");
		delegacionCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		if (Boolean.TRUE.equals(input.getDelRegPat()) && b) {
			Cell delegacionOtrasCell = rowDelegacion.createCell(8);
			delegacionOtrasCell.setCellValue("Otras Delegaciones");
			delegacionOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		} else {
			Cell delegacionOtrasCell = rowDelegacion.createCell(7);
			delegacionOtrasCell.setCellValue("Otras Delegaciones");
			delegacionOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		}
		if (Boolean.TRUE.equals(input.getDelRegPat()) && b) {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D14:H14"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("I14:L14"));
		} else {
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("D14:G14"));
			sheetNacional.addMergedRegion(CellRangeAddress.valueOf("H14:K14"));
		}

		Row rowEncabezadoNacional = sheetNacional.createRow(14);
		Cell tipoArchivoCell = rowEncabezadoNacional.createCell(1);
		tipoArchivoCell.setCellValue("Identificador de archivo");
		tipoArchivoCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		if (Boolean.TRUE.equals(input.getDelRegPat())) {

			Cell delDetCell = rowEncabezadoNacional.createCell(2);
			delDetCell.setCellValue("Delegacion");
			delDetCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
					HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

			headerReportDel(rowEncabezadoNacional, font, workbook);
			
		} else {
						
			headerReportNal(rowEncabezadoNacional, font, workbook);

		}
	}

	private CellStyle createStyle(XSSFFont font, HorizontalAlignment hAlign, VerticalAlignment vAlign, short cellColor,
			boolean cellBorder, short cellBorderColor, Workbook workbook, boolean wrap) {

		CellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setFillForegroundColor(cellColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(hAlign);
		style.setVerticalAlignment(vAlign);
		style.setWrapText(wrap);

		if (cellBorder) {
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);

			style.setTopBorderColor(cellBorderColor);
			style.setLeftBorderColor(cellBorderColor);
			style.setRightBorderColor(cellBorderColor);
			style.setBottomBorderColor(cellBorderColor);
		}

		return style;
	}
	
	private void fDetailReport(Row rowTotal, CifrasControlDTO cifrasControlDTO, CellStyle style) {

		Cell totalDetCell = rowTotal.createCell(2);
		totalDetCell.setCellValue(cifrasControlDTO.getNumTotalRegistros());
		totalDetCell.setCellStyle(style);

		Cell correctosDetCell = rowTotal.createCell(3);
		correctosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosCorrectos());
		correctosDetCell.setCellStyle(style);

		Cell erroneosDetCell = rowTotal.createCell(4);
		erroneosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosError());
		erroneosDetCell.setCellStyle(style);

		Cell duplicadosDetCell = rowTotal.createCell(5);
		duplicadosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosDup());
		duplicadosDetCell.setCellStyle(style);

		Cell susAjusDetCell = rowTotal.createCell(6);
		susAjusDetCell.setCellValue(cifrasControlDTO.getNumRegistrosSus());
		susAjusDetCell.setCellStyle(style);

		Cell correctosOtrasDetCell = rowTotal.createCell(7);
		correctosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosCorrectosOtras());
		correctosOtrasDetCell.setCellStyle(style);

		Cell erroneosOtrasDetCell = rowTotal.createCell(8);
		erroneosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosErrorOtras());
		erroneosOtrasDetCell.setCellStyle(style);

		Cell duplicadosOtrasDetCell = rowTotal.createCell(9);
		duplicadosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosDupOtras());
		duplicadosOtrasDetCell.setCellStyle(style);

		Cell susAjusOtrasDetCell = rowTotal.createCell(10);
		susAjusOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosSusOtras());
		susAjusOtrasDetCell.setCellStyle(style);
	}
	
	private void fDetailReportDel(Row rowTotal, CifrasControlDTO cifrasControlDTO, CellStyle style) {

		Cell totalDetCell = rowTotal.createCell(3);
		totalDetCell.setCellValue(cifrasControlDTO.getNumTotalRegistros());
		totalDetCell.setCellStyle(style);

		Cell correctosDetCell = rowTotal.createCell(4);
		correctosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosCorrectos());
		correctosDetCell.setCellStyle(style);

		Cell erroneosDetCell = rowTotal.createCell(5);
		erroneosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosError());
		erroneosDetCell.setCellStyle(style);

		Cell duplicadosDetCell = rowTotal.createCell(6);
		duplicadosDetCell.setCellValue(cifrasControlDTO.getNumRegistrosDup());
		duplicadosDetCell.setCellStyle(style);

		Cell susAjusDetCell = rowTotal.createCell(7);
		susAjusDetCell.setCellValue(cifrasControlDTO.getNumRegistrosSus());
		susAjusDetCell.setCellStyle(style);

		Cell correctosOtrasDetCell = rowTotal.createCell(8);
		correctosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosCorrectosOtras());
		correctosOtrasDetCell.setCellStyle(style);

		Cell erroneosOtrasDetCell = rowTotal.createCell(9);
		erroneosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosErrorOtras());
		erroneosOtrasDetCell.setCellStyle(style);

		Cell duplicadosOtrasDetCell = rowTotal.createCell(10);
		duplicadosOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosDupOtras());
		duplicadosOtrasDetCell.setCellStyle(style);

		Cell susAjusOtrasDetCell = rowTotal.createCell(11);
		susAjusOtrasDetCell.setCellValue(cifrasControlDTO.getNumRegistrosSusOtras());
		susAjusOtrasDetCell.setCellStyle(style);
	}
	
	private void headerReportNal(Row rowEncabezadoNacional, XSSFFont font, Workbook workbook) {
		
		Cell totalCell = rowEncabezadoNacional.createCell(2);
		totalCell.setCellValue("Total");
		totalCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
		
		Cell correctosCell = rowEncabezadoNacional.createCell(3);
		correctosCell.setCellValue(CORRECTOS);
		correctosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell erroneosCell = rowEncabezadoNacional.createCell(4);
		erroneosCell.setCellValue(ERRONEOS);
		erroneosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell duplicadosCell = rowEncabezadoNacional.createCell(5);
		duplicadosCell.setCellValue(DUPLICADOS);
		duplicadosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell susAjusCell = rowEncabezadoNacional.createCell(6);
		susAjusCell.setCellValue(SUS_AJUSTE);
		susAjusCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell correctosOtrasCell = rowEncabezadoNacional.createCell(7);
		correctosOtrasCell.setCellValue(CORRECTOS);
		correctosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell erroneosOtrasCell = rowEncabezadoNacional.createCell(8);
		erroneosOtrasCell.setCellValue(ERRONEOS);
		erroneosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell duplicadosOtrasCell = rowEncabezadoNacional.createCell(9);
		duplicadosOtrasCell.setCellValue(DUPLICADOS);
		duplicadosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell susAjusOtrasCell = rowEncabezadoNacional.createCell(10);
		susAjusOtrasCell.setCellValue(SUS_AJUSTE);
		susAjusOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
	}
	
	private void headerReportDel(Row rowEncabezadoNacional, XSSFFont font, Workbook workbook) {
		
		Cell totalesCell = rowEncabezadoNacional.createCell(3);
		totalesCell.setCellValue("Total");
		totalesCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
			HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));	
	
		Cell correctosCell = rowEncabezadoNacional.createCell(4);
		correctosCell.setCellValue(CORRECTOS);
		correctosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell erroneosCell = rowEncabezadoNacional.createCell(5);
		erroneosCell.setCellValue(ERRONEOS);
		erroneosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell duplicadosCell = rowEncabezadoNacional.createCell(6);
		duplicadosCell.setCellValue(DUPLICADOS);
		duplicadosCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell susAjusCell = rowEncabezadoNacional.createCell(7);
		susAjusCell.setCellValue(SUS_AJUSTE);
		susAjusCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell correctosOtrasCell = rowEncabezadoNacional.createCell(8);
		correctosOtrasCell.setCellValue(CORRECTOS);
		correctosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell erroneosOtrasCell = rowEncabezadoNacional.createCell(9);
		erroneosOtrasCell.setCellValue(ERRONEOS);
		erroneosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell duplicadosOtrasCell = rowEncabezadoNacional.createCell(10);
		duplicadosOtrasCell.setCellValue(DUPLICADOS);
		duplicadosOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));

		Cell susAjusOtrasCell = rowEncabezadoNacional.createCell(11);
		susAjusOtrasCell.setCellValue(SUS_AJUSTE);
		susAjusOtrasCell.setCellStyle(createStyle(font, HorizontalAlignment.CENTER, VerticalAlignment.CENTER,
				HSSFColor.GREY_25_PERCENT.index, false, HSSFColor.WHITE.index, workbook, true));
	}

}
