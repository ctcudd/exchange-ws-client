/**
 * 
 */
package com.microsoft.exchange.config;

import javax.inject.Inject;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.ImportResource;

import com.microsoft.exchange.ExchangeWebServices;
import com.microsoft.exchange.impl.BaseExchangeCalendarDataDao;

/**
 * @author ctcudd
 *
 */
@Configuration
//@EnableRetry

//@ImportResource({"classpath:/com/microsoft/exchange/exchangeContext-usingImpersonation.xml"})
@ImportResource({"classpath:/test-contexts/exchangeContext.xml"})
public class ImpersonationConfig {
//	
//	@Bean
//	BaseExchangeCalendarDataDao baseExchangeCalendarDataDao(){
//		return new BaseExchangeCalendarDataDao();
//	}
//	
//	@Inject 
//	ExchangeWebServices exchangeWebServices;
}
