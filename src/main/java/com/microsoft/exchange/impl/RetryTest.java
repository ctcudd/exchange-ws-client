/**
 * 
 */
package com.microsoft.exchange.impl;

import static org.junit.Assert.*;

import java.util.ArrayList;
import java.util.Collection;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.mockito.Matchers;
import org.mockito.Mockito;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;
import org.springframework.test.context.support.AnnotationConfigContextLoader;

import com.microsoft.exchange.ExchangeWebServices;
import com.microsoft.exchange.config.TestConfig;
import com.microsoft.exchange.messages.DeleteItem;
import com.microsoft.exchange.types.ItemIdType;

/**
 * @author Collin Cudd
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(loader=AnnotationConfigContextLoader.class,classes=TestConfig.class)
public class RetryTest {

	@Autowired
	ExchangeWebServices client;
	
	@Autowired
	ExchangeCalendarDataDao exchangeCalendarDataDao;

	@Test
	public void isAutowired() {
		assertNotNull(exchangeCalendarDataDao);
		assertNotNull(client);		
	}	
	
	@Test
	public void testRetry() throws Exception {
		ItemIdType itemId = new ItemIdType();
		itemId.setId("fakeitemid");
		Collection<ItemIdType> itemIds = new ArrayList<ItemIdType>();
		itemIds.add(itemId);
		
		boolean deleted = exchangeCalendarDataDao.deleteCalendarItemsWithRetry("upn", itemIds);
		Mockito.verify(client, Mockito.times(3)).deleteItem(Matchers.any(DeleteItem.class));
		assertTrue(deleted);
	}
}
