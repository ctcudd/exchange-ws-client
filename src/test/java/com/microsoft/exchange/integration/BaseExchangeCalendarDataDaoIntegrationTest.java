/**
 * See the NOTICE file distributed with this work
 * for additional information regarding copyright ownership.
 * Board of Regents of the University of Wisconsin System
 * licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file
 * except in compliance with the License. You may obtain a
 * copy of the License at:
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on
 * an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package com.microsoft.exchange.integration;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import junit.framework.Assert;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.junit.Before;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.dao.support.DataAccessUtils;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.ibm.icu.util.TimeZone;
import com.microsoft.exchange.exception.ExchangeRuntimeException;
import com.microsoft.exchange.impl.BaseExchangeCalendarDataDao;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.messages.FindItemResponse;
import com.microsoft.exchange.messages.GetServerTimeZones;
import com.microsoft.exchange.types.BaseFolderType;
import com.microsoft.exchange.types.BodyTypeType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.CalendarItemTypeType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.TimeZoneDefinitionType;


@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(locations= {"classpath:test-contexts/exchangeContext.xml"})
public class BaseExchangeCalendarDataDaoIntegrationTest {
	protected final Log log = LogFactory.getLog(this.getClass());

	String username = "someusername";
	
	@Value("${integration.email:someemailaddress@on.yourexchangeserver.edu}")
	protected String upn;
	
	@Autowired
	protected BaseExchangeCalendarDataDao exchangeCalendarDataDao;
	
	@Before
	public void setup(){
				//"Test1-DoIT@aims.wisc.edu";
	}

	@Rule
	public ExpectedException exception = ExpectedException.none();
	
	@Test
	public void isAutowired() {
		assertNotNull(exchangeCalendarDataDao);
		assertNotNull(upn);
	}
	
	@Test
	public void resolveEmailAddresses() {
		Set<String> results = exchangeCalendarDataDao.resolveEmailAddresses(upn);
		assertNotNull(results);
		for(String r: results) log.info(r);
	}
	
	@Test
	public void resolveUpn() {
		String resolved = exchangeCalendarDataDao.resolveUpn(upn);
		assertNotNull(resolved);
		log.info(upn+" resolved to "+ resolved);
	}
	
	/**
	 * Executes a {@link GetServerTimeZones} request, retrieving all {@link TimeZoneDefinitionType}s from Exchange.
	 * Attempts to map each {@link TimeZoneDefinitionType} to an {@link TimeZone} and logs the result.
	 */
	@Test
	public void getTimeZoneIds(){
		int validTimeZoneCount = 0;
		Set<TimeZoneDefinitionType> zones = exchangeCalendarDataDao.getServerTimeZones(null, false);
		log.info("Found "+zones.size() +" exchange time zones.");
		for(TimeZoneDefinitionType zone: zones){
			String sysTimeZoneID = TimeZone.getIDForWindowsID(zone.getId(), "US");
			
			if(StringUtils.isNotBlank(sysTimeZoneID)){
				log.info(zone.getId() +" mapped to "+ sysTimeZoneID);
				
				String windowsID = TimeZone.getWindowsID(sysTimeZoneID);
				assertEquals(zone.getId(), windowsID);
				validTimeZoneCount++;
			}else{
				log.warn("no mapping for windowsID="+zone.getId());
			}
			
		}
		log.info("Succesfully mapped "+validTimeZoneCount+"/"+zones.size()+" WindowsTimeZones");
	}
	
	@Test
	public void getPrimaryCalendarFolder(){
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		assertNotNull(primaryCalendarFolder);
	}
	
	@Test
	public void getPrimaryTaskFolder(){
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		assertNotNull(primaryCalendarFolder);
	}
	
	@Test
	public void getCalendarFolders(){
		Map<String, String> folderMap = exchangeCalendarDataDao.getCalendarFolderMap(upn);
		for(String folderId: folderMap.keySet()){
			String folderName = folderMap.get(folderId);
			log.info(folderName);
		}
	}
	
	@Test
	public void getCalendarByDateRangeCalendarViewFindItem(){
		Date startDate  = new Date();
		Date endDate = DateUtils.addDays(startDate, 1);
		
		Set<ItemIdType> itemIds = exchangeCalendarDataDao.findCalendarItemIds(upn, startDate, endDate);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, itemIds);
		Assert.assertEquals(itemIds.size(), calendarItems.size());
		log.info("Found "+calendarItems.size()+" calendarItems");
		for(CalendarItemType c :  calendarItems){
			log.info("subject: "+c.getSubject());
			log.info("uid: "+c.getUID());
			log.info("recurrenceId: "+c.getRecurrenceId().toString());
			log.info("created: "+c.getDateTimeCreated().toString());
			log.info("");
		}
	}
	
	@Test
	public void purgeCancelledCalendarItems(){
		upn ="cclose@wisc.edu";
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		boolean purged = exchangeCalendarDataDao.purgeCancelledCalendarItems(upn, primaryCalendarFolder.getFolderId() );
		assertTrue(purged);
	}
	
	@Test
	public void getCancelledCalendarItemsByDateRangePageViewFindItem(){
		upn ="cclose@wisc.edu";
		Date now  = new Date();
		Date startDate = DateUtils.addDays(now, -100);
		Date endDate = DateUtils.addDays(now, 100);
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		assertNotNull(primaryCalendarFolder);
		Integer totalCount = primaryCalendarFolder.getTotalCount();
		log.info("primaryCalendarFolder.totalCount="+totalCount);
		FindItem request = exchangeCalendarDataDao.getRequestFactory().constructIndexedPageViewFindItemCancelledIdsByDateRange(startDate, endDate, Collections.singleton(primaryCalendarFolder.getFolderId()));
		FindItemResponse response = exchangeCalendarDataDao.getWebServices().findItem(request);
		Pair<Set<ItemIdType>, Integer> parsed = exchangeCalendarDataDao.getResponseUtils().parseFindItemIdResponse(response);
		Integer nextOffset = parsed.getRight();
		
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, parsed.getLeft());
		log.info("Found "+calendarItems.size()+" calendarItems");
		for(CalendarItemType c :  calendarItems){
			if(!c.isIsCancelled()){
				log.info("fail");
			}
		}
		assertTrue(exchangeCalendarDataDao.deleteCalendarItems(upn,  parsed.getLeft()));
		while(nextOffset > 0){
			request = exchangeCalendarDataDao.getRequestFactory().constructIndexedPageViewFindItemCancelledIdsByDateRange(nextOffset, startDate, endDate, Collections.singleton(primaryCalendarFolder.getFolderId()));
		    response = exchangeCalendarDataDao.getWebServices().findItem(request);
		    parsed = exchangeCalendarDataDao.getResponseUtils().parseFindItemIdResponse(response);
			nextOffset = parsed.getRight();
			
			calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, parsed.getLeft());
			log.info("Found "+calendarItems.size()+" calendarItems");
			for(CalendarItemType c :  calendarItems){
				assertTrue(c.isIsCancelled());
			}
			assertTrue(exchangeCalendarDataDao.deleteCalendarItems(upn,  parsed.getLeft()));
		}

		

	}
	
	@Test
	public void getCalendarByDateRangePageViewFindItem(){

		Date now  = new Date();
		Date startDate = DateUtils.addDays(now, -100);
		Date endDate = DateUtils.addDays(now, 100);
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		assertNotNull(primaryCalendarFolder);
		Integer totalCount = primaryCalendarFolder.getTotalCount();
		log.info("primaryCalendarFolder.totalCount="+totalCount);
		FindItem request = exchangeCalendarDataDao.getRequestFactory().constructIndexedPageViewFindItemIdsByDateRange(startDate, endDate, Collections.singleton(primaryCalendarFolder.getFolderId()));
		FindItemResponse response = exchangeCalendarDataDao.getWebServices().findItem(request);
		Pair<Set<ItemIdType>, Integer> parsed = exchangeCalendarDataDao.getResponseUtils().parseFindItemIdResponse(response);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, parsed.getLeft());
		log.info("Found "+calendarItems.size()+" calendarItems");
		Collection<ItemIdType> cancelled = new HashSet<ItemIdType>();
		for(CalendarItemType c :  calendarItems){
			log.info(c.getSubject() +" - "+c.getCalendarItemType());
			log.info("cancelled: "+c.isIsCancelled());
			log.info("created:   "+c.getDateTimeCreated().toString());
			log.info("start : "+ c.getStart().toString());
			log.info("end	: "+ c.getEnd().toString());
			if(c.getRecurrenceId() != null){
				log.info(c.getRecurrence());
			}
			log.info(c.getUID());
			log.info("");
			if(c.isIsCancelled()){
				cancelled.add(c.getItemId());
			}
		}
		log.info("Found "+cancelled.size()+" cancelled calendarItems");
		boolean deleted = exchangeCalendarDataDao.deleteCalendarItems(upn, cancelled);
		assertTrue(deleted);
		
		
	}
	
	@Test
	public void getAllCalendarItems(){
		BaseFolderType primaryCalendarFolder = exchangeCalendarDataDao.getPrimaryCalendarFolder(upn);
		Set<FolderIdType> singleton = Collections.singleton(primaryCalendarFolder.getFolderId());
		Set<ItemIdType> itemIds = exchangeCalendarDataDao.findAllItemIds(upn, singleton);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, itemIds);
		log.info("Found "+calendarItems.size()+" calendarItems");
		for(CalendarItemType c :  calendarItems){
			log.info("subject: "+c.getSubject());
			log.info("uid: "+c.getUID());
			log.info("recurrenceId: "+c.getRecurrenceId().toString());
			log.info("created: "+c.getDateTimeCreated().toString());
			log.info("");
		}
	}
	
	@Test
	public void getTaskFolders(){
		Map<String, String> folderMap = exchangeCalendarDataDao.getTaskFolderMap(upn);
		for(String folderId: folderMap.keySet()){
			String folderName = folderMap.get(folderId);
			log.info(folderName);
		}
	}
	
	@Test
	public void deleteTaskFolders(){
		Map<String, String> taskFoldersMap = exchangeCalendarDataDao.getTaskFolderMap(upn);
		for(String taskId :  taskFoldersMap.keySet()){
			String taskFolderName = taskFoldersMap.get(taskId);
			if(taskFolderName.equals("Tasks")) continue;
			FolderIdType taskFolderId= new FolderIdType();
			taskFolderId.setId(taskId);
			
			log.info("deleting taskFolder '"+taskFolderName+"' ");
			boolean deleteTaskFolderSuccess = exchangeCalendarDataDao.deleteFolder(upn, taskFolderId);
			assertTrue(deleteTaskFolderSuccess);
		}
	}
	
	@Test
	public void deleteCalendarFolders(){
		Map<String, String> calFolders = exchangeCalendarDataDao.getCalendarFolderMap(upn);
		for(String calId :  calFolders.keySet()){
			String calName = calFolders.get(calId);

			if(calName.startsWith("blspycha@wisc.edu")){
				FolderIdType calFolderId= new FolderIdType();
				calFolderId.setId(calId);
				
				log.info("deleting calendarFolder named: '"+calName+"' ");
				boolean deleteCalendarFolderSuccess = exchangeCalendarDataDao.deleteCalendarFolder(upn, calFolderId);
				assertTrue(deleteCalendarFolderSuccess);
			}
		}
	}
	
	@Test
	public void purgeEmptyCalendarFolders(){
		Map<String, String> calFolders = exchangeCalendarDataDao.getCalendarFolderMap(upn);
		for(String calId :  calFolders.keySet()){
			String calName = calFolders.get(calId);

			if(calName.startsWith("blspycha@wisc.edu")){
				FolderIdType calFolderId= new FolderIdType();
				calFolderId.setId(calId);
				
				log.info("checking calendarFolder named: '"+calName+"' ");
				boolean deleteCalendarFolderSuccess = exchangeCalendarDataDao.deleteEmptyCalendarFolder(upn, calFolderId);
				if(deleteCalendarFolderSuccess){
					log.info(" '"+calName+"' was deleted");
				}else{
					log.info(" '"+calName+"' was NOT deleted");
				}
			}
		}
	}
	
	/**
	 * This method can fail when there are many (1k+) items in a folder, probably due to throttling limits....
	 * "Exchange Web Services are not currently available for this request because none of the Client Access Servers in the destination site could process the request." 
	 *
	 *This method can also fail when disposalType is set to MOVE_TO_DELETED_ITEMS
	 *If impersonation is used you have to use the MoveToDeletedItems method...
	 *see note here: http://msdn.microsoft.com/en-us/library/office/aa580484(v=exchg.150).aspx
	 *
	 *
	 */
	@Test
	public void deleteFolder(){
		FolderIdType calendarFolderId = exchangeCalendarDataDao.getCalendarFolderId(upn, "A2");
		assertNotNull(calendarFolderId);
		
		boolean deleteFolderResult = exchangeCalendarDataDao.deleteFolder(upn, calendarFolderId);
		assertTrue(deleteFolderResult);
	}
	
	@Test
	public void findAllItemIds(){
		FolderIdType calendarFolderId = exchangeCalendarDataDao.getCalendarFolderId(upn, "A2");
		assertNotNull(calendarFolderId);
		
		Set<FolderIdType> singleton = Collections.singleton(calendarFolderId);
		Set<ItemIdType> itemIds = exchangeCalendarDataDao.findAllItemIds(upn, singleton);
		
		for(ItemIdType itemId :itemIds){
			log.info(itemId.getId());
		}
	}
	
	@Test
	public void emptyCalendarFolder(){
		FolderIdType calendarFolderId = exchangeCalendarDataDao.getCalendarFolderId(upn, "A1");
		assertNotNull(calendarFolderId);
		boolean emptyCalendarFolder = exchangeCalendarDataDao.emptyCalendarFolder(upn, calendarFolderId);
		assertTrue(emptyCalendarFolder);
		
		Set<ItemIdType> findItemIds = exchangeCalendarDataDao.findItemIds(upn, Collections.singleton(calendarFolderId));
		assertTrue(findItemIds.isEmpty());
	}
	
	/**
	 * An error response that includes the ErrorCannotDeleteObject error
	 * code will be returned for a DeleteItem operation when a delegate
	 * tries to delete an item in the principal's mailbox by setting the
	 * DisposalType to MoveToDeletedItems. To delete an item by moving it to
	 * the Deleted Items folder, a delegate must use the MoveItem operation.
	 * 
	 * @see http://msdn.microsoft.com/en-us/library/office/aa580484(v=exchg.150).aspx
	 */
	@Test
	public void createDeleteCalendarFolder(){
		String displayName = "TEST "+upn;
		FolderIdType folderId = exchangeCalendarDataDao.createCalendarFolder(upn, displayName);
		assertNotNull(folderId);
		
		boolean deleteFolderSuccess = false;
		
//		try{
//			deleteFolderSuccess =exchangeCalendarDataDao.deleteFolder(upn, DisposalType.MOVE_TO_DELETED_ITEMS, folderId);
//			fail("MOVE_TO_DELETED_ITEMS should have thrown an exception!");
//		}catch(ExchangeCannotDeleteRuntimeException e){	}
//		
		assertFalse(deleteFolderSuccess);
		log.info("deleteFolder via MOVE_TO_DELETED_ITEMS failed as expected, attempting SOFT_DELETE");
		deleteFolderSuccess = exchangeCalendarDataDao.deleteFolder(upn, folderId);
		if(deleteFolderSuccess){
			log.info("deleteFolder via SOFT_DELETE success!");
		}
//		else{
//			log.info("deleteFolder via SOFT_DELETE failure, attempting HARD_DELETE");
//			deleteFolderSuccess = exchangeCalendarDataDao.deleteFolder(upn, DisposalType.HARD_DELETE, folderId);
//		}
		assertTrue(deleteFolderSuccess);
		exception.expect(ExchangeRuntimeException.class);
		folderId = exchangeCalendarDataDao.getCalendarFolderId(upn, displayName);
	}
	
	@Test
	public void getEmptyDeleteCalendarFolder(){
		String displayName = "TEST "+upn;
		FolderIdType calendarFolderId = exchangeCalendarDataDao.getCalendarFolderId(upn, displayName);
		assertNotNull(calendarFolderId);
		
		boolean result = exchangeCalendarDataDao.deleteCalendarFolder(upn, calendarFolderId);
		assertTrue(result);
		
		exception.expect(ExchangeRuntimeException.class);
		calendarFolderId = exchangeCalendarDataDao.getCalendarFolderId(upn, displayName);
		
		assertNull(calendarFolderId);
	}
	

	@Test
	public void getCalendarItemById(){
		String itemIdString = "AAMkADIxZTFmMjM1LTBmYjEtNDg0MS1iNTQ0LWU3OTI1ZWZhZDlkOABGAAAAAACN05QVBWdrSarKY7LX12MEBwDaH0wtNHE9T69RMmiPkcoJAAAAAAEOAADaH0wtNHE9T69RMmiPkcoJAAC4vbyDAAA=";
		ItemIdType itemId = new ItemIdType();
		itemId.setId(itemIdString);
		Collection<CalendarItemType> calendarItems = exchangeCalendarDataDao.getCalendarItems(upn, Collections.singleton(itemId));
	}
	
	@Test
	public void createGetDeleteCalendarItem(){
		UUID seed = UUID.randomUUID();
		CalendarItemType calendarItem = new CalendarItemType();
		
		calendarItem.setUID(seed.toString());
		ItemIdType calendarItemId = exchangeCalendarDataDao.createCalendarItem(upn, calendarItem);
		assertNotNull(calendarItemId);
		
		Collection<CalendarItemType> createdCalendarItems = exchangeCalendarDataDao.getCalendarItems(upn, Collections.singleton(calendarItemId));
		CalendarItemType createdCalendarItem = DataAccessUtils.singleResult(createdCalendarItems);
		assertNotNull(createdCalendarItem);
		assertNotNull(createdCalendarItem.getStart());
		assertEquals(seed.toString(), calendarItem.getUID());
		
		boolean deleteSuccess = exchangeCalendarDataDao.deleteCalendarItems(upn, Collections.singleton(calendarItemId));
		assertTrue(deleteSuccess);
	}
	
	@Test
	public void sendEmail(){
		List<String> recips = new ArrayList<String>();
		recips.add(upn);
		
		String replyTo = upn;
		String subject ="this is subject";
		String messageBody = ";this is body;";
		BodyTypeType bodyType = BodyTypeType.TEXT;
		FolderIdType folderIdType = null;
 		ItemIdType emailItemId = exchangeCalendarDataDao.createEmailMessage(recips, replyTo, subject, messageBody, bodyType, folderIdType);
	}
}
