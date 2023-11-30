package mx.gob.imss.cit.mspmccifrascontrol.services.impl;

import java.util.List;
import mx.gob.imss.cit.mspmccommons.integration.model.CifrasControlMovimientosResponseDTO;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import mx.gob.imss.cit.mspmccifrascontrol.MsPmcCifrasControlInput;
import mx.gob.imss.cit.mspmccifrascontrol.integration.dao.MsPmcCifrasControlRepository;
import mx.gob.imss.cit.mspmccifrascontrol.services.MsPmcCifrasControlService;
import mx.gob.imss.cit.mspmccommons.exception.BusinessException;

@Service("msPmcCifrasControlService")
public class MsPmcCifrasControlSerivceImpl implements MsPmcCifrasControlService {

	@Autowired
	private MsPmcCifrasControlRepository msPmcCifrasControlRepository;

	@Override
	public List<CifrasControlMovimientosResponseDTO> getCifrasControl(MsPmcCifrasControlInput input) throws BusinessException {
		return msPmcCifrasControlRepository.getCifrasControl(input);
	}

}
