/**
 * 
 */
package com.microsoft.exchange.config;

import javax.inject.Inject;
import javax.xml.bind.JAXBContext;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.env.Environment;
import org.springframework.oxm.jaxb.Jaxb2Marshaller;
import org.springframework.ws.transport.http.HttpComponentsMessageSender;

import com.microsoft.exchange.impl.http.CustomHttpComponentsMessageSender;

/**
 * @author ctcudd
 *
 */
@Configuration
@PropertySource("classpath:/exchange.properties")
public class ImpersonationConfig {
	
	@Inject
	Environment env;
	
	@Bean
	JAXBContext jaxb2Context(){
		return jaxb2Marshaller().getJaxbContext();
	}
	
	@Bean
	Jaxb2Marshaller jaxb2Marshaller(){
		Jaxb2Marshaller marshaller = new Jaxb2Marshaller();
		String[] contextPaths = { 
			"com.microsoft.exchange.messages",
			"com.microsoft.exchange.types"
				};
		marshaller.setContextPaths(contextPaths);
		return marshaller;
	}
	
//	@Bean
//	ExchangeWebServicesClient exchangeWebServicesClient(){
//		
//		Jaxb2Marshaller marshaller = jaxb2Marshaller();
//		
//		RequestServerTimeZoneInterceptor tzInterceptor = new RequestServerTimeZoneInterceptor();
//		tzInterceptor.setJaxbContext(marshaller.getJaxbContext());
//		
//		RequestServerVersionClientInterceptor versionInterceptor = new RequestServerVersionClientInterceptor();
//		versionInterceptor.setJaxbContext(marshaller.getJaxbContext());
//		
//		ExchangeImpersonationClientInterceptor impersonationInterceptor = new ExchangeImpersonationClientInterceptor();
//		impersonationInterceptor.setJaxbContext(marshaller.getJaxbContext());
//		
//		ClientInterceptor[] interceptors = {tzInterceptor, versionInterceptor, impersonationInterceptor};
//		ExchangeWebServicesClient client = new ExchangeWebServicesClient();
//		client.setMarshaller(marshaller);
//		client.setUnmarshaller(marshaller);
//		client.setMessageFactory(new SaajSoapMessageFactory());
//		
//		client.setInterceptors(interceptors);
//		client.setMessageSender(httpComponentsMessageSender());
//		client.setDefaultUri(env.getRequiredProperty("endpoint"));
//		
//		String truststorePath = env.getRequiredProperty("truststore");
//		Resource resource = new ClassPathResource(truststorePath);
//		client.setTrustStore(resource);
//
//		return client;
//	}
	
	@Bean
	HttpComponentsMessageSender httpComponentsMessageSender(){
		CustomHttpComponentsMessageSender sender = new CustomHttpComponentsMessageSender();
		
		sender.setMaxTotalConnections(
				env.getProperty("http.maxTotalConnections", int.class, 10));
		
		sender.setDefaultMaxPerRouteOverride(
				env.getProperty("http.maxConnectionsPerRoute", Integer.class, new Integer(10)));
		
		sender.setConnectionTimeout(
				env.getProperty("http.connectionTimeout", int.class, 1200000));
		
		sender.setReadTimeout(
				env.getProperty("http.readTimeout", int.class, 1200000));
	
		sender.setPreemptiveAuthEnabled(
				env.getProperty("http.preemptiveAuthEnabled", boolean.class, false));
		
		sender.setNtlmAuthEnabled(
				env.getProperty("http.ntlmAuthEnabled", boolean.class, false));

		return sender;
	}
}
