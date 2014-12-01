/**
 * 
 */
package com.microsoft.exchange.config;

import javax.xml.bind.JAXBContext;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.oxm.jaxb.Jaxb2Marshaller;


/**
 * @author ctcudd
 *
 */
@Configuration
public class JaxbConfig {

	@Bean
	public JAXBContext jaxbContext(){
		JAXBContext context = JAXBContext.
	}
	
	@Bean
	public Jaxb2Marshaller jaxb2Marshaller() {
		Jaxb2Marshaller marshaller = new Jaxb2Marshaller();
		marshaller.setContextPaths("com.microsoft.exchange.messages",
									"com.microsoft.exchange.types",
									"com.microsoft.exchange.autodiscover");
		return marshaller;
	}
}

