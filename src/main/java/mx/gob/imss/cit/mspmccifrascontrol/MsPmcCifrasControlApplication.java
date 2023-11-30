package mx.gob.imss.cit.mspmccifrascontrol;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.client.RestTemplateBuilder;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpMethod;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.builders.WebSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.annotation.web.configuration.WebSecurityConfigurerAdapter;
import org.springframework.security.web.authentication.UsernamePasswordAuthenticationFilter;
import org.springframework.web.client.RestTemplate;

import mx.gob.imss.cit.mspmccifrascontrol.security.JWTAuthorizationFilter;
import mx.gob.imss.cit.mspmccifrascontrol.security.service.TokenValidateService;

@SpringBootApplication
public class MsPmcCifrasControlApplication {

	public static void main(String[] args) {
		SpringApplication.run(MsPmcCifrasControlApplication.class, args);
	}

	@Bean
	public RestTemplate restTemplate(RestTemplateBuilder restTemplateBuilder) {
		return restTemplateBuilder.build();
	}

	@EnableWebSecurity
	@Configuration
	class WebSecurityConfig extends WebSecurityConfigurerAdapter {

		@Bean
		public TokenValidateService tokenPmcValidateService() {
			return new TokenValidateService();
		}

		@Override
		protected void configure(HttpSecurity http) throws Exception {
			http.csrf().disable()
					.addFilterAfter(new JWTAuthorizationFilter(tokenPmcValidateService()),
							UsernamePasswordAuthenticationFilter.class)
					.authorizeRequests().antMatchers(HttpMethod.POST, "/cifrascontrol/v1/cifrascontrol**").permitAll()
					.antMatchers(HttpMethod.POST, "/cifrascontrol/v1/reportpdf").permitAll()
					.antMatchers(HttpMethod.POST, "/cifrascontrol/v1/report**").permitAll().anyRequest()
					.authenticated();
		}
		
		@Override
		public void configure(WebSecurity webSecurity) {
			webSecurity.ignoring().antMatchers(
					"/swagger-resources/**",
					"/swagger-ui.html",
					"/v2/api-docs",
					"/webjars/**"
			);
		}
	}
}
