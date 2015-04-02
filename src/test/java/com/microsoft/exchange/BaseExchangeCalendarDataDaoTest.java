/**
 * 
 */
package com.microsoft.exchange;

import javax.xml.bind.JAXBContext;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.junit.Assert;
import org.junit.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;

import com.microsoft.exchange.impl.BaseExchangeCalendarDataDao;

/**
 * Tests for {@link BaseExchangeCalendarDataDao}.
 * 
 * @author Collin Cudd
 *
 */
public class BaseExchangeCalendarDataDaoTest {

	@Mock JAXBContext jaxbContext;
	@Mock ExchangeWebServices ExchangeWebServices;
	@InjectMocks BaseExchangeCalendarDataDao calendarDao = new BaseExchangeCalendarDataDao();
	protected final Log log = LogFactory.getLog(this.getClass());

	/**
	 * Control test for {@link BaseExchangeCalendarDataDao#getWaitTimeExp(int)}
	 */
	@Test
	public void getWaitTimeExp_control(){
		long waitTimeMillis = 0L;
		for(int i=0; i< calendarDao.getMaxRetries(); i++){
			waitTimeMillis = BaseExchangeCalendarDataDao.getWaitTimeExp(i);
			double l = ((double) Math.pow(2, i) * 1000d);
			Assert.assertTrue(waitTimeMillis > l);
			double waitTimeSeconds = waitTimeMillis / 1000d;
			log.info("retryCount="+i+" waitTime="+String.format("%.2f",waitTimeSeconds)+"(s)");
		}
	}
}
