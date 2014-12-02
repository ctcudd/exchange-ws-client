/**
 * 
 */
package com.microsoft.exchange;

import static org.junit.Assert.*;

import javax.xml.bind.JAXBElement;

import org.junit.Test;

import com.microsoft.exchange.types.BasePathToElementType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.PathToUnindexedFieldType;
import com.microsoft.exchange.types.SetItemFieldType;
import com.microsoft.exchange.types.UnindexedFieldURIType;

/**
 * @author ctcudd
 *
 */
public class ExchangeRequestFactoryTest {
	ExchangeRequestFactory requestFactory = new ExchangeRequestFactory();
	
	@Test
	public void constructSetCalendarItemLegacyFreeBusy(){
		CalendarItemType c = new CalendarItemType();
		SetItemFieldType setFreeBusy = requestFactory.constructSetCalendarItemLegacyFreeBusy(c);
		assertNotNull(setFreeBusy);
		JAXBElement<? extends BasePathToElementType> path = setFreeBusy.getPath();
		assertNotNull(path);
		BasePathToElementType basePath = path.getValue();
		assertNotNull(basePath);
		assertTrue(basePath instanceof PathToUnindexedFieldType);
		PathToUnindexedFieldType unindexedPath = (PathToUnindexedFieldType) basePath;
		assertNotNull(unindexedPath);
		assertEquals(UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS, unindexedPath.getFieldURI());
	}
}
