package mx.gob.imss.cit.mspmccifrascontrol.integration.dao.impl;

import java.util.*;
import java.util.stream.Collectors;
import mx.gob.imss.cit.mspmccommons.enums.EstadoRegistroEnum;
import mx.gob.imss.cit.mspmccommons.integration.model.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.mongodb.core.MongoOperations;
import org.springframework.data.mongodb.core.aggregation.*;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.stereotype.Repository;
import io.micrometer.core.instrument.util.StringUtils;
import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccifrascontrol.integration.dao.MsPmcCifrasControlRepository;
import mx.gob.imss.cit.mspmccifrascontrol.integration.dao.ParametroRepository;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;
import mx.gob.imss.cit.mspmccommons.utils.AggregationUtils;
import mx.gob.imss.cit.mspmccommons.utils.CustomAggregationOperation;
import mx.gob.imss.cit.mspmccommons.utils.DateUtils;

@Repository
public class MsPmcCifrasControlRepositoryImpl implements MsPmcCifrasControlRepository {

	@Autowired
	private MongoOperations mongoOperations;

	@Autowired
	private ParametroRepository parametroRepository;

	private final Logger logger = LoggerFactory.getLogger(getClass());

	@Override
	public List<CifrasControlMovimientosResponseDTO> getCifrasControl(MsPmcCifrasControlInput input) throws BusinessException {

		List<CifrasControlMovimientosResponseDTO> response;

		// Se calculan las fechas inicio y fin para la consulta
		Date fecProcesoIni = DateUtils.calculateBeginDate(input.getFromYear(), input.getFromMonth(), null);
		Date fecProcesoFin = DateUtils.calculateEndDate(input.getToYear(), input.getToMonth(), null);
		Criteria cFecProcesoCarga = null;
		Criteria cFecProcesoCargaDel = null;

		if (fecProcesoIni != null && fecProcesoFin != null) {
			cFecProcesoCarga = new Criteria().andOperator(Criteria.where("fecProcesoCarga").gt(fecProcesoIni),
					Criteria.where("fecProcesoCarga").lte(fecProcesoFin));
			cFecProcesoCargaDel = Criteria.where("aseguradoDTO.fecAlta").gt(fecProcesoIni).lte(fecProcesoFin);
		}
		logger.info("cveDelegacion recibida: {}", input.getCveDelegation());

		Criteria cCveEstadoArchivo = Criteria.where("cveEstadoArchivo").is("2");
		Criteria cCveOrigenArchivo = null;

		if (StringUtils.isNotBlank(input.getCveTipoArchivo()) && StringUtils.isNotEmpty(input.getCveTipoArchivo())) {
			cCveOrigenArchivo = Criteria.where("cveOrigenArchivo").is(input.getCveTipoArchivo());
		}

		logger.info("cveDelegacion recibida: {}", input.getCveDelegation());

		logger.info("--------------Query de agregacion-------------------");
		if (input.getDelRegPat() == null || !input.getDelRegPat()) {
			TypedAggregation<ArchivoDTO> aggregation = buildAggregation(cFecProcesoCarga, cCveEstadoArchivo, cCveOrigenArchivo);
			logger.info("Agregacion: {}", aggregation);
			AggregationResults<ArchivoDTO> aggregationResult = mongoOperations.aggregate(aggregation, ArchivoDTO.class);
			List<ArchivoDTO> listArchivos = aggregationResult.getMappedResults();
			response = listArchivos.stream().map(archivoDTO -> {
				CifrasControlMovimientosResponseDTO cifras = new CifrasControlMovimientosResponseDTO();
				cifras.setCveOrigenArchivo(archivoDTO.getCveOrigenArchivo());
				cifras.setCorrecto(archivoDTO.getCifrasControlDTO().getNumRegistrosCorrectos());
				cifras.setErroneo(archivoDTO.getCifrasControlDTO().getNumRegistrosError());
				cifras.setDuplicado(archivoDTO.getCifrasControlDTO().getNumRegistrosDup());
				cifras.setSusceptible(archivoDTO.getCifrasControlDTO().getNumRegistrosSus());
				cifras.setCorrectoOtras(archivoDTO.getCifrasControlDTO().getNumRegistrosCorrectosOtras());
				cifras.setErroneoOtras(archivoDTO.getCifrasControlDTO().getNumRegistrosErrorOtras());
				cifras.setDuplicadoOtras(archivoDTO.getCifrasControlDTO().getNumRegistrosDupOtras());
				cifras.setSusceptibleOtras(archivoDTO.getCifrasControlDTO().getNumRegistrosSusOtras());
				cifras.setTotal(archivoDTO.getCifrasControlDTO().getNumTotalRegistros());
				return cifras;
			}).collect(Collectors.toList());
		} else {
			TypedAggregation<CifrasControlMovimientosResponseDTO> aggregationDel = buildAggregationDel(
					cFecProcesoCargaDel, cCveOrigenArchivo);
			logger.info("AgregacionDelegacion: {}", aggregationDel);
			AggregationResults<CifrasControlMovimientosResponseDTO> aggregationResultDetalle = mongoOperations
					.aggregate(aggregationDel, DetalleRegistroDTO.class, CifrasControlMovimientosResponseDTO.class);
			response = aggregationResultDetalle.getMappedResults();
		}
			logger.info("----------------------------------------------------");

		return response;
	}
	
	private TypedAggregation<CifrasControlMovimientosResponseDTO> buildAggregationDel(Criteria cFecProcesoCarga, Criteria cCveOrigenArchivo) {
		String addFieldsDelNssJson = buildAddFieldsString("desDelNss", buildCondString(Boolean.FALSE));
		logger.info(addFieldsDelNssJson);
		CustomAggregationOperation addFieldsDelNss = new CustomAggregationOperation(addFieldsDelNssJson);
		String addFieldsDelPatJson = buildAddFieldsString("desDelPat", buildCondString(Boolean.TRUE));
		logger.info(addFieldsDelPatJson);
		CustomAggregationOperation addFieldsDelPat = new CustomAggregationOperation(addFieldsDelPatJson);
		String groupDelNssJson = buildGroupString("desDelNss");
		logger.info(groupDelNssJson);
		CustomAggregationOperation groupDelNss = new CustomAggregationOperation(groupDelNssJson);
		String groupDelPatJson = buildGroupString("desDelPat");
		logger.info(groupDelPatJson);
		CustomAggregationOperation groupDelPat = new CustomAggregationOperation(groupDelPatJson);
		String groupSumJson = buildFinalGroupString();
		logger.info(groupSumJson);
		CustomAggregationOperation groupSum = new CustomAggregationOperation(groupSumJson);
		ProjectionOperation projection = Aggregation.project()
				.andExpression("concatArrays(delNss, delRegPat)")
				.as("conteos");

		UnwindOperation unwind = Aggregation.unwind("conteos");
		String sortJson = "{ $sort: { _id: 1 }}";
		CustomAggregationOperation sort = new CustomAggregationOperation(sortJson);
		FacetOperation facet = Aggregation.facet().and(groupDelNss).as("delNss")
				.and(groupDelPat).as("delRegPat");

		List<AggregationOperation> aggregationOperationList = Arrays.asList(
				AggregationUtils.validateMatchOp(cFecProcesoCarga),
				AggregationUtils.validateMatchOp(cCveOrigenArchivo),
				addFieldsDelNss,
				addFieldsDelPat,
				facet,
				projection,
				unwind,
				groupSum,
				AggregationUtils.validateMatchOp(Criteria.where("_id").ne(null)),
				sort
				);
		aggregationOperationList = aggregationOperationList.stream()
				.filter(Objects::nonNull)
				.collect(Collectors.toList());
		return Aggregation.newAggregation(CifrasControlMovimientosResponseDTO.class, aggregationOperationList);
	}

	private String buildAddFieldsString(String name, String cond) {
		String addFields = "{ $addFields: {";
		addFields = addFields.concat(name)
				.concat(": {")
				.concat(cond)
				.concat("} } }");
		return addFields;
	}

	private String buildCondString(boolean delPat) {
		String cond = "$cond: { if: { ";
		cond = delPat ?
				cond.concat("$or:[{$in:['$aseguradoDTO.cveEstadoRegistro',[5,8,7]]}," +
						"{$eq:['$aseguradoDTO.desDelegacionNss','$patronDTO.desDelRegPatronal']}] },")
					.concat("then: null, else: '$patronDTO.desDelRegPatronal' }") :
				cond.concat("$in: ['$aseguradoDTO.cveEstadoRegistro', [5,8,7]] },")
					.concat("then: '$patronDTO.desDelRegPatronal', else: '$aseguradoDTO.desDelegacionNss' }");
		return cond;
	}

	private String buildGroupString(String by) {
		String group = "{ $group: { _id: '$";
		group = group.concat(by).concat("', cveOrigenArchivo: { $first: '$cveOrigenArchivo' },");
		for(EstadoRegistroEnum estadoRegistro : EstadoRegistroEnum.values()) {
			if (estadoRegistro.getCveEstadoRegistro() >= 1 && estadoRegistro.getCveEstadoRegistro() <= 8) {
				group = group.concat(formatName(estadoRegistro.name()))
						.concat(": { $sum: { $cond: [{$eq: ['$aseguradoDTO.cveEstadoRegistro', ")
						.concat(String.valueOf(estadoRegistro.getCveEstadoRegistro()))
						.concat("]}, { $sum: 1 }, { $sum: 0 }] } },");
			}
		}
		group = group.concat("total: { $sum: { $cond: [{$in: ['$aseguradoDTO.cveEstadoRegistro', [1,2,3,4,5,6,7,8]]}, { $sum: 1 }, { $sum: 0 }] } } } }");
		return group;
	}

	private String buildFinalGroupString() {
		String group = "{ $group: { _id: '$conteos._id', cveOrigenArchivo: { $first: '$conteos.cveOrigenArchivo' },";
		for(EstadoRegistroEnum estadoRegistro : EstadoRegistroEnum.values()) {
			if (estadoRegistro.getCveEstadoRegistro() >= 1 && estadoRegistro.getCveEstadoRegistro() <= 8) {
				String formattedName = formatName(estadoRegistro.name());
				group = group.concat(formattedName)
						.concat(": { $sum: '$conteos.")
						.concat(formattedName)
						.concat("' },");
			}
		}
		group = group.concat("total: { $sum: '$conteos.total' } } }");
		return group;
	}

	private String formatName(String name) {
		StringBuilder stringBuilder = new StringBuilder(name.toLowerCase());
		for (int i = 0; i < stringBuilder.length(); i++) {
			if (stringBuilder.charAt(i) == '_') {
				stringBuilder.deleteCharAt(i);
				stringBuilder.replace(i, i + 1, String.valueOf(Character.toUpperCase(stringBuilder.charAt(i))));
			}
		}
		return stringBuilder.toString();
	}
	
	private TypedAggregation<ArchivoDTO> buildAggregation(Criteria cFecProcesoCarga, Criteria cCveEstadoArchivo,
			Criteria cCveOrigenArchivo) {
		String jsonOpperation = "{ $project: {"
				+ "        'objectIdArchivo': 1,"
				+ "        'cifrasControlDTO': 1,"
				+ "        'cveOrigenArchivo': 1"
				+ "    }}";
		CustomAggregationOperation projection = new CustomAggregationOperation(jsonOpperation);
		List<AggregationOperation> aggregationOperationList = Arrays.asList(
				AggregationUtils.validateMatchOp(cFecProcesoCarga),
				AggregationUtils.validateMatchOp(cCveEstadoArchivo),
				AggregationUtils.validateMatchOp(cCveOrigenArchivo),
				projection
				);
		aggregationOperationList = aggregationOperationList.stream()
				.filter(Objects::nonNull)
				.collect(Collectors.toList());
		return Aggregation.newAggregation(ArchivoDTO.class, aggregationOperationList);
	}

}
