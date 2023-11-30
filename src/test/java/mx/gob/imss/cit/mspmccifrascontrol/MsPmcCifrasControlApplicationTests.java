package mx.gob.imss.cit.mspmccifrascontrol;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import mx.gob.imss.cit.mspmccifrascontrol.services.ReporteService;

@SpringBootTest
class MsPmcCifrasControlApplicationTests {
	
	@Autowired
	ReporteService reporte;

	@Test
	void contextLoads() {
		assertNotNull(reporte);
	}

}
