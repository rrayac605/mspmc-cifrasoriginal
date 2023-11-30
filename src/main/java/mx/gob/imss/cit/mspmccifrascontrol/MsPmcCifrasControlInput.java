package mx.gob.imss.cit.mspmccifrascontrol;

import lombok.Data;

@Data
public class MsPmcCifrasControlInput {

	private String fromMonth;

	private String fromYear;

	private String toMonth;

	private String toYear;

	private String cveTipoArchivo;

	private String cveDelegation;

	private Boolean delRegPat;

	private Boolean isPdfReport;

	private Integer page;

	private String desDelegation;

}
