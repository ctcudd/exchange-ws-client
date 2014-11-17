/**
 * 
 */
package com.microsoft.exchange.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.util.Collection;
import java.util.Collections;
import java.util.HashSet;
import java.util.Set;

import org.junit.Test;
import org.springframework.dao.support.DataAccessUtils;

import com.microsoft.exchange.messages.UpdateItem;
import com.microsoft.exchange.messages.UpdateItemResponse;
import com.microsoft.exchange.types.AttendeeType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.EmailAddressType;
import com.microsoft.exchange.types.ImportanceChoicesType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.NonEmptyArrayOfAttendeesType;
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangesType;
import com.microsoft.exchange.types.SetItemFieldType;


/**
 * @author ctcudd
 *
 */
public class UpdateCalendarItemIntegrationTest extends BaseExchangeCalendarDataDaoIntegrationTest {

	@Test
	public void isAutowired(){
		assertNotNull(exchangeCalendarDataDao);
		assertNotNull(upn);
	}
	
	@Test
	public void updateCalendarItemSubject_control(){
		String initialSubject = "THISISSUBJECTYES";
		String newSubject ="newSubject";
		ImportanceChoicesType initialImportance = ImportanceChoicesType.HIGH;
		
		CalendarItemType c = new CalendarItemType();
		c.setSubject(initialSubject);
		c.setImportance(initialImportance);
		
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> createdItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(createdItems);
		assertEquals(1, createdItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(createdItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		assertEquals(initialSubject, createdItem.getSubject());
		assertEquals(initialImportance, createdItem.getImportance());
		
		createdItem.setSubject(newSubject);
		SetItemFieldType setItem = exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemSubject(createdItem);
		NonEmptyArrayOfItemChangesType changes = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItemChanges(createdItem, Collections.singleton(setItem));
		
		UpdateItem request = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItem(createdItem, changes);
		UpdateItemResponse response = exchangeCalendarDataDao.getWebServices().updateItem(request);
		Set<ItemIdType> updatedItemIds = exchangeCalendarDataDao.getResponseUtils().parseUpdateItemResponse(response);
		
		Collection<CalendarItemType> updatedItems = exchangeCalendarDataDao.getCalendarItems(upn, updatedItemIds);
		
		//ensures a singular non null result
		assertNotNull(updatedItems);
		assertEquals(1, updatedItems.size());
		CalendarItemType updatedItem = DataAccessUtils.singleResult(updatedItems);
		
		//make sure the ids match
		assertEquals(newSubject, updatedItem.getSubject());
		assertEquals(initialImportance, updatedItem.getImportance());
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, Collections.singleton(updatedItem.getItemId()));
		assertTrue(deleteSuccess);
	}
	
	@Test
	public void updateCalendarItemAddAttendee_control(){
		String initialSubject = "THISISSUBJECTYES";
		ImportanceChoicesType initialImportance = ImportanceChoicesType.HIGH;
		
		CalendarItemType c = new CalendarItemType();
		c.setSubject(initialSubject);
		c.setImportance(initialImportance);
		
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> createdItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(createdItems);
		assertEquals(1, createdItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(createdItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		assertEquals(initialSubject, createdItem.getSubject());
		assertEquals(initialImportance, createdItem.getImportance());
		
		AttendeeType newAttendee = new AttendeeType();
		EmailAddressType mailbox = new EmailAddressType();
		mailbox.setEmailAddress("me@test.com");
		newAttendee.setMailbox(mailbox);
		
		NonEmptyArrayOfAttendeesType attendees = new NonEmptyArrayOfAttendeesType();
		attendees.getAttendees().add(newAttendee);
		
		createdItem.setRequiredAttendees(attendees);
		SetItemFieldType updateField = exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemRequiredAttendees(createdItem);
		NonEmptyArrayOfItemChangesType updateChanges = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItemChanges(createdItem, Collections.singleton(updateField));
		UpdateItem request = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItem(createdItem, updateChanges);
		
		UpdateItemResponse response = exchangeCalendarDataDao.getWebServices().updateItem(request);
		Set<ItemIdType> updatedItemIds = exchangeCalendarDataDao.getResponseUtils().parseUpdateItemResponse(response);
		
		Collection<CalendarItemType> updatedItems = exchangeCalendarDataDao.getCalendarItems(upn, updatedItemIds);
		
		//ensures a singular non null result
		assertNotNull(updatedItems);
		assertEquals(1, updatedItems.size());
		CalendarItemType updatedItem = DataAccessUtils.singleResult(updatedItems);
		
		//make sure the ids match
		assertEquals(initialSubject, updatedItem.getSubject());
		assertEquals(initialImportance, updatedItem.getImportance());
		assertEquals(1, updatedItem.getRequiredAttendees().getAttendees().size());
		
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, Collections.singleton(updatedItem.getItemId()));
		assertTrue(deleteSuccess);
	}
	
	@Test
	public void updateCalendarItem_AddAttendee_ChangeSubject(){
		String initialSubject = "THISISSUBJECTYES";
		String changeSubject = "newsubjecthere";
		ImportanceChoicesType initialImportance = ImportanceChoicesType.HIGH;
		
		CalendarItemType c = new CalendarItemType();
		c.setSubject(initialSubject);
		c.setImportance(initialImportance);
		
		ItemIdType createdItemId = exchangeCalendarDataDao.createCalendarItem(upn, c);
		assertNotNull(createdItemId);
		
		//retrieve the newly creatd item
		Set<ItemIdType> expectedItemId = Collections.singleton(createdItemId);
		Collection<CalendarItemType> createdItems = exchangeCalendarDataDao.getCalendarItems(upn, expectedItemId);
		
		//ensures a singular non null result
		assertNotNull(createdItems);
		assertEquals(1, createdItems.size());
		CalendarItemType createdItem = DataAccessUtils.singleResult(createdItems);
		
		//make sure the ids match
		assertEquals(createdItemId, createdItem.getItemId());
		assertEquals(initialSubject, createdItem.getSubject());
		assertEquals(initialImportance, createdItem.getImportance());
		assertEquals(null, createdItem.getRequiredAttendees());
		
		AttendeeType newAttendee = new AttendeeType();
		EmailAddressType mailbox = new EmailAddressType();
		mailbox.setEmailAddress("me@test.com");
		newAttendee.setMailbox(mailbox);
		
		NonEmptyArrayOfAttendeesType attendees = new NonEmptyArrayOfAttendeesType();
		attendees.getAttendees().add(newAttendee);
		
		//change the calendar event
		createdItem.setRequiredAttendees(attendees);
		createdItem.setSubject(changeSubject);
		
		//produce the changes
		Collection<SetItemFieldType> changes = new HashSet<SetItemFieldType>();
		changes.add(exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemRequiredAttendees(createdItem));
		changes.add(exchangeCalendarDataDao.getRequestFactory().constructSetCalendarItemSubject(createdItem));
		
		//construct the request
		NonEmptyArrayOfItemChangesType updateChanges = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItemChanges(createdItem, changes);
		UpdateItem request = exchangeCalendarDataDao.getRequestFactory().constructUpdateCalendarItem(createdItem, updateChanges);
		
		UpdateItemResponse response = exchangeCalendarDataDao.getWebServices().updateItem(request);
		Set<ItemIdType> updatedItemIds = exchangeCalendarDataDao.getResponseUtils().parseUpdateItemResponse(response);
		
		Collection<CalendarItemType> updatedItems = exchangeCalendarDataDao.getCalendarItems(upn, updatedItemIds);
		
		//ensures a singular non null result
		assertNotNull(updatedItems);
		assertEquals(1, updatedItems.size());
		CalendarItemType updatedItem = DataAccessUtils.singleResult(updatedItems);
		
		//make sure the changes occured
		assertEquals(changeSubject, updatedItem.getSubject());
		assertEquals(initialImportance, updatedItem.getImportance());
		assertEquals(1, updatedItem.getRequiredAttendees().getAttendees().size());
		
		//cleanup
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, Collections.singleton(updatedItem.getItemId()));
		assertTrue(deleteSuccess);
	}
	
}
