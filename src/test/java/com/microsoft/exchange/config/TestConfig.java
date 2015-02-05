/**
 * 
 */
package com.microsoft.exchange.config;

import javax.inject.Inject;
import javax.xml.bind.JAXBElement;

import org.mockito.Matchers;
import org.mockito.Mockito;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.EnableAspectJAutoProxy;
import org.springframework.context.annotation.Import;
import org.springframework.core.env.Environment;
import org.springframework.oxm.jaxb.Jaxb2Marshaller;
import org.springframework.retry.annotation.EnableRetry;
import org.springframework.ws.client.core.WebServiceTemplate;
import org.springframework.ws.client.support.interceptor.ClientInterceptor;
import org.springframework.ws.soap.saaj.SaajSoapMessageFactory;
import org.springframework.ws.transport.http.HttpComponentsMessageSender;

import com.microsoft.exchange.exception.ExchangeWebServicesRuntimeException;
import com.microsoft.exchange.impl.BaseExchangeCalendarDataDao;
import com.microsoft.exchange.impl.ExchangeCalendarDataDao;
import com.microsoft.exchange.impl.ExchangeImpersonationClientInterceptor;
import com.microsoft.exchange.impl.ExchangeWebServicesClient;
import com.microsoft.exchange.impl.RequestServerTimeZoneInterceptor;
import com.microsoft.exchange.impl.RequestServerVersionClientInterceptor;
import com.microsoft.exchange.messages.ArrayOfResponseMessagesType;
import com.microsoft.exchange.messages.DeleteItem;
import com.microsoft.exchange.messages.DeleteItemResponse;
import com.microsoft.exchange.messages.ObjectFactory;
import com.microsoft.exchange.messages.ResponseCodeType;
import com.microsoft.exchange.messages.ResponseMessageType;
import com.microsoft.exchange.types.ResponseClassType;

/**
 * @author ctcudd
 *
 */
@Configuration
//@EnableAspectJAutoProxy(proxyTargetClass = true)
@EnableRetry
@Import(value=ImpersonationConfig.class)
public class TestConfig {
	
	@Inject
	Environment env;
	
	@Inject 
	Jaxb2Marshaller marshaller;
	
	@Inject
	HttpComponentsMessageSender messageSender;
	
	@Bean
	WebServiceTemplate webServiceTemplate(){
		WebServiceTemplate template = new WebServiceTemplate();
		
		template.setDefaultUri(env.getRequiredProperty("endpoint"));
		template.setInterceptors(clientInterceptors());
		template.setMarshaller(marshaller);
		template.setUnmarshaller(marshaller);
		template.setMessageFactory(new SaajSoapMessageFactory());
		template.setMessageSender(messageSender);
		
		return template;
	}
	
	@Bean
	ClientInterceptor[] clientInterceptors(){
		RequestServerTimeZoneInterceptor tzInterceptor = new RequestServerTimeZoneInterceptor();
		tzInterceptor.setJaxbContext(marshaller.getJaxbContext());
		
		RequestServerVersionClientInterceptor versionInterceptor = new RequestServerVersionClientInterceptor();
		versionInterceptor.setJaxbContext(marshaller.getJaxbContext());
		
		ExchangeImpersonationClientInterceptor impersonationInterceptor = new ExchangeImpersonationClientInterceptor();
		impersonationInterceptor.setJaxbContext(marshaller.getJaxbContext());
		
		ClientInterceptor[] interceptors = {tzInterceptor, versionInterceptor, impersonationInterceptor};
		return interceptors;
	}
	
	@Bean
	ExchangeWebServicesClient exchangeWebServicesClient(){
		ExchangeWebServicesClient client = Mockito.mock(ExchangeWebServicesClient.class);
		client.setWebServiceTemplate(webServiceTemplate());
		
		ResponseMessageType responseMessageType = new ResponseMessageType();
		responseMessageType.setResponseClass(ResponseClassType.SUCCESS);
		responseMessageType.setResponseCode(ResponseCodeType.NO_ERROR);
	
		ObjectFactory of = new ObjectFactory();
		JAXBElement<ResponseMessageType> element = of.createArrayOfResponseMessagesTypeDeleteItemResponseMessage(responseMessageType);
		
		ArrayOfResponseMessagesType array = new ArrayOfResponseMessagesType();
		array.getCreateItemResponseMessagesAndDeleteItemResponseMessagesAndGetItemResponseMessages().add(element);
		
		DeleteItemResponse response = new DeleteItemResponse();
		response.setResponseMessages(array);
		
		Mockito.when(client.deleteItem(Matchers.any(DeleteItem.class)))
			.thenThrow(ExchangeWebServicesRuntimeException.class)
			.thenThrow(ExchangeWebServicesRuntimeException.class)
			.thenReturn(response);
		
		return client;
	}
	
	@Bean
	ExchangeCalendarDataDao exchangeCalendarDataDao(){
		BaseExchangeCalendarDataDao baseExchangeCalendarDataDao = new BaseExchangeCalendarDataDao();
		baseExchangeCalendarDataDao.setWebServices(exchangeWebServicesClient());
		baseExchangeCalendarDataDao.setJaxbContext(marshaller.getJaxbContext());
		return baseExchangeCalendarDataDao;
	}
}
