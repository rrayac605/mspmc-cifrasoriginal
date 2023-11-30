package mx.gob.imss.cit.mspmccifrascontrol.integration.dao;

import java.util.List;
import mx.gob.imss.cit.mspmccommons.integration.model.CifrasControlMovimientosResponseDTO;

import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;

public interface MsPmcCifrasControlRepository {

	List<CifrasControlMovimientosResponseDTO> getCifrasControl(MsPmcCifrasControlInput input) throws BusinessException;

}
