package mx.gob.imss.cit.mspmccifrascontrol.security;

import java.io.IOException;
import java.util.logging.Logger;
import javax.servlet.FilterChain;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.web.filter.OncePerRequestFilter;

import mx.gob.imss.cit.mspmccifrascontrol.security.service.TokenValidateService;

public class JWTAuthorizationFilter extends OncePerRequestFilter {

	protected final Logger log = Logger.getLogger(getClass().getName());

	protected TokenValidateService tokenValidateService;

	public JWTAuthorizationFilter(TokenValidateService tokenValidateService) {
		super();
		this.tokenValidateService = tokenValidateService;
	}

	@Override
	protected void doFilterInternal(HttpServletRequest request, HttpServletResponse response, FilterChain chain)
			throws ServletException, IOException {

		chain.doFilter(request, response);
	}

}
