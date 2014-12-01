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
package com.microsoft.exchange.impl;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.bind.JAXBContext;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.Validate;
import org.apache.commons.lang.time.StopWatch;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.commons.validator.routines.EmailValidator;
import org.joda.time.Interval;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.dao.support.DataAccessUtils;
import org.springframework.util.CollectionUtils;

import com.microsoft.exchange.ExchangeDateUtils;
import com.microsoft.exchange.ExchangeRequestFactory;
import com.microsoft.exchange.ExchangeResponseUtils;
import com.microsoft.exchange.ExchangeWebServices;
import com.microsoft.exchange.exception.ExchangeExceededFindCountLimitRuntimeException;
import com.microsoft.exchange.exception.ExchangeInvalidUPNRuntimeException;
import com.microsoft.exchange.exception.ExchangeRuntimeException;
import com.microsoft.exchange.exception.ExchangeWebServicesRuntimeException;
import com.microsoft.exchange.messages.CreateFolder;
import com.microsoft.exchange.messages.CreateFolderResponse;
import com.microsoft.exchange.messages.CreateItem;
import com.microsoft.exchange.messages.CreateItemResponse;
import com.microsoft.exchange.messages.DeleteFolder;
import com.microsoft.exchange.messages.DeleteFolderResponse;
import com.microsoft.exchange.messages.DeleteItem;
import com.microsoft.exchange.messages.DeleteItemResponse;
import com.microsoft.exchange.messages.EmptyFolder;
import com.microsoft.exchange.messages.EmptyFolderResponse;
import com.microsoft.exchange.messages.FindFolder;
import com.microsoft.exchange.messages.FindFolderResponse;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.messages.FindItemResponse;
import com.microsoft.exchange.messages.GetFolder;
import com.microsoft.exchange.messages.GetFolderResponse;
import com.microsoft.exchange.messages.GetItem;
import com.microsoft.exchange.messages.GetItemResponse;
import com.microsoft.exchange.messages.GetServerTimeZones;
import com.microsoft.exchange.messages.GetServerTimeZonesResponse;
import com.microsoft.exchange.messages.ResolveNames;
import com.microsoft.exchange.messages.ResolveNamesResponse;
import com.microsoft.exchange.messages.UpdateItem;
import com.microsoft.exchange.messages.UpdateItemResponse;
import com.microsoft.exchange.types.ArrayOfRecipientsType;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.BaseFolderType;
import com.microsoft.exchange.types.BodyType;
import com.microsoft.exchange.types.BodyTypeType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemCreateOrDeleteOperationType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.ConnectingSIDType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.EmailAddressType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.FolderQueryTraversalType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MessageType;
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangesType;
import com.microsoft.exchange.types.SetItemFieldType;
import com.microsoft.exchange.types.SingleRecipientType;
import com.microsoft.exchange.types.TaskType;
import com.microsoft.exchange.types.TasksFolderType;
import com.microsoft.exchange.types.TimeZoneDefinitionType;


/**
 * TODO - what is the intent for this class? Does it belong within the client?
 * If so, is the intent that it be a base class for interacting with {@link ExchangeWebServices}?
 * Should it be abstract? 
 * 
 * @author Collin Cudd
 */
public class BaseExchangeCalendarDataDao {

    //================================================================================
    // Properties 
    //================================================================================
	protected final Log log = LogFactory.getLog(this.getClass());
	
	private JAXBContext jaxbContext;
	private ExchangeWebServices webServices;
	
	// TODO no references, needed internally?
	private ExchangeRequestFactory requestFactory = new ExchangeRequestFactory();
	
	// TODO no references, needed internally?
	private ExchangeResponseUtils responseUtils = new ExchangeResponseUtilsImpl();

	private int maxRetries = 10;
	
	@Value("${username}")
	private String adminUpn;
	
	@Value("${admin.sendas}")
	private String adminSendAs;
	
    //================================================================================
    // Getters/Setters 
    //================================================================================
	public String getAdminSendAs(){
		return this.adminSendAs;
	}
	
	public String getAdminUpn(){
		return this.adminUpn;
	}
	
	public int getMaxRetries() {
		return maxRetries;
	}
	public void setMaxRetries(int maxRetries) {
		this.maxRetries = maxRetries;
	}
	public ExchangeWebServices getWebServices() {
		return webServices;
	}
	/**
	 * @param exchangeWebServices the exchangeWebServices to set
	 */
	@Autowired @Qualifier("ewsClient")
	public void setWebServices(ExchangeWebServices exchangeWebServices) {
		this.webServices = exchangeWebServices;
	}

	/**
	 * @return the requestFactory
	 */
	public ExchangeRequestFactory getRequestFactory() {
		return requestFactory;
	}
	/**
	 * @param requestFactory the requestFactory to set
	 */
	@Autowired(required=false)
	public void setRequestFactory(ExchangeRequestFactory exchangeRequestFactory) {
		this.requestFactory = exchangeRequestFactory;
	}
	
	public ExchangeResponseUtils getResponseUtils() {
		return responseUtils;
	}
	public void setResponseUtils(ExchangeResponseUtils exchangeResponseUtils) {
		this.responseUtils = exchangeResponseUtils;
	}
	
	/**
	 * @return the jaxbContext
	 */
	public JAXBContext getJaxbContext() {
		return jaxbContext;
	}
	/**
	 * @param jaxbContext the jaxbContext to set
	 */
	@Autowired
	public void setJaxbContext(JAXBContext jaxbContext) {
		this.jaxbContext = jaxbContext;
	}
	
	/**
	 * Function for determining how long to sleep before retrying a failed operation
	 * @param retryCount
	 * @return
	 */
	public static long getWaitTimeExp(int retryCount) {
		long waitTime = ((long) Math.pow(2, retryCount) * 100L);
	    return waitTime;
	}
	
	protected void setContextCredentials(String upn) {
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		ConnectingSIDType connectingSID = new ConnectingSIDType();
		connectingSID.setPrincipalName(upn);
		ThreadLocalImpersonationConnectingSIDSourceImpl.setConnectingSID(connectingSID);
	}
	
	public CalendarFolderType getCalendarFolder(String upn, FolderIdType folderId) {
		BaseFolderType folder = getFolder(upn, folderId);
		CalendarFolderType calendarFolderType = null;
		if(folder instanceof CalendarFolderType) {
			calendarFolderType=  (CalendarFolderType) folder;
		}
		return calendarFolderType;
	}
	
	public TasksFolderType getTaskFolder(String upn, FolderIdType folderId) {
		BaseFolderType folder = getFolder(upn, folderId);
		TasksFolderType taskFolderType = null;
		if(folder instanceof TasksFolderType) {
			taskFolderType =  (TasksFolderType) folder;
		}
		return taskFolderType;
	}
	
	public BaseFolderType getFolder(String upn, FolderIdType folderIdType){
		setContextCredentials(upn);
		GetFolder getFolderRequest = getRequestFactory().constructGetFolderById(folderIdType);
		GetFolderResponse getFolderResponse = getWebServices().getFolder(getFolderRequest);
		Set<BaseFolderType> response = getResponseUtils().parseGetFolderResponse(getFolderResponse);
		return DataAccessUtils.singleResult(response);
	}
	
	protected BaseFolderType getPrimaryFolder(String upn, DistinguishedFolderIdNameType parent) {
		setContextCredentials(upn);
		GetFolder getFolderRequest = getRequestFactory().constructGetFolderByDistinguishedName(parent);
		GetFolderResponse getFolderResponse = getWebServices().getFolder(getFolderRequest);
		Set<BaseFolderType> response = getResponseUtils().parseGetFolderResponse(getFolderResponse);
		return DataAccessUtils.singleResult(response);
	}
	
	public BaseFolderType getPrimaryCalendarFolder(String upn){
		return getPrimaryFolder(upn, DistinguishedFolderIdNameType.CALENDAR);		
	}
	
	public BaseFolderType getPrimaryTaskFolder(String upn) {
		return getPrimaryFolder(upn, DistinguishedFolderIdNameType.TASKS);
	}
	
	private Set<BaseFolderType> getFoldersByType(String upn, DistinguishedFolderIdNameType parent){
		Set<BaseFolderType> folders = new HashSet<BaseFolderType>();
		BaseFolderType baseFolderType = getPrimaryFolder(upn, parent);
		if(null != baseFolderType) {
			folders.add(baseFolderType);
		}
		Set<BaseFolderType> seondaryFolders = getSeondaryFolders(upn, parent);
		if(!CollectionUtils.isEmpty(seondaryFolders)) {
			for(BaseFolderType b: seondaryFolders ) {
				if(baseFolderType.getClass().equals(b.getClass())) {
					folders.add(b);
				}
			}
		}
		return folders;
	}
	
	/**
	 * Gets all secondary folders for given DistinguisheFolderName
	 * @param upn
	 * @param parent
	 * @return
	 */
	private Set<BaseFolderType> getSeondaryFolders(String upn, DistinguishedFolderIdNameType parent) {
		Validate.notEmpty(upn, "upn cannnot be empty");
		setContextCredentials(upn);
		FindFolder findFolderRequest = getRequestFactory().constructFindFolder(parent, DefaultShapeNamesType.ALL_PROPERTIES, FolderQueryTraversalType.DEEP);
		FindFolderResponse findFolderResponse = getWebServices().findFolder(findFolderRequest);
		return getResponseUtils().parseFindFolderResponse(findFolderResponse);
	}
	/*
	 * return all calendar folders
	 */
	public Set<BaseFolderType> getAllCalendarFolders(String upn) {
		Validate.notEmpty(upn, "upn cannnot be empty");
		return getFoldersByType(upn, DistinguishedFolderIdNameType.CALENDAR);
	}
	
	public Set<BaseFolderType> getAllTaskFolders(String upn) {
		Validate.notEmpty(upn, "upn cannnot be empty");
		return getFoldersByType(upn, DistinguishedFolderIdNameType.TASKS);
	}
	
	public FolderIdType getCalendarFolderId(String upn, String calendarName) {
		Map<String, String> calendarFolderMap = getCalendarFolderMap(upn);
		if(!CollectionUtils.isEmpty(calendarFolderMap) && calendarFolderMap.containsValue(calendarName)) {
			for(String c_id: calendarFolderMap.keySet()) {
				String c_name = calendarFolderMap.get(c_id);
				if(calendarName.equals(c_name)) {
					FolderIdType folderIdType = new FolderIdType();
					folderIdType.setId(c_id);
					return folderIdType;
				}
			}
		}
		throw new ExchangeRuntimeException("No calendar folder with name of '"+calendarName+"' for "+upn); 
	}

	public Map<String, String> getCalendarFolderMap(String upn){
		Map<String, String> calendarsMap = new HashMap<String, String>();
		Set<BaseFolderType> allCalendarFolders = getAllCalendarFolders(upn);
		for(BaseFolderType folderType: allCalendarFolders) {
			String name = folderType.getDisplayName();
			String id = folderType.getFolderId().getId();
			calendarsMap.put(id, name);
		}
		return calendarsMap;
	}
	
	public Map<String, String> getTaskFolderMap(String upn){
		Map<String, String> taskFolderMap = new HashMap<String, String>();
		Set<BaseFolderType> allTaskFolders = getAllTaskFolders(upn);
		for(BaseFolderType b : allTaskFolders) {
			String displayName = b.getDisplayName();
			String id = b.getFolderId().getId();
			taskFolderMap.put( id,displayName);
		}
		return taskFolderMap;
	}
	
	public Set<ItemIdType> findCalendarItemIds(String upn, Date startDate, Date endDate){
		return findCalendarItemIdsInternal(upn, startDate, endDate, null, 0);
	}
	public Set<ItemIdType> findCalendarItemIds(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarIds) {
		return findCalendarItemIdsInternal(upn, startDate, endDate, calendarIds, 0);
	}
	
	
	
	
	/**
	 * This method uses a CalendarView...
	 * 
	 * The FindItem operation can return results in a CalendarView element. The
	 * CalendarView element returns single calendar items and all occurrences.
	 * If a CalendarView element is not used, single calendar items and
	 * recurring master calendar items are returned. The occurrences must be
	 * expanded from the recurring master if a CalendarView element is not used.
	 * 
	 * -- http://msdn.microsoft.com/en-us/library/office/aa566107(v=exchg.140).
	 * aspx
	 * 
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @param calendarIds
	 * @param depth
	 * @return
	 */
	private Set<ItemIdType> findCalendarItemIdsInternal(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarIds, int depth){
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notNull(startDate, "startDate argument cannot be null");
		Validate.notNull(endDate, "endDate argument cannot be null");
		
		//folderIds can be null

		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("findCalendarItemIdsInternal(upn="+upn+",startDate="+startDate+",+endDate="+endDate+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			FindItem request = getRequestFactory().constructCalendarViewFindCalendarItemIdsByDateRange(startDate, endDate, calendarIds);
			try {
				FindItemResponse response = getWebServices().findItem(request);
				return getResponseUtils().parseFindItemIdResponseNoOffset(response);
				
			}catch(ExchangeInvalidUPNRuntimeException e0) {
				log.warn("findCalendarItemIdsInternal(upn="+upn+",startDate="+startDate+",+endDate="+endDate+",...) ExchangeInvalidUPNRuntimeException.  Attempting to resolve valid upn... - failure #"+newDepth);
				
				String resolvedUpn = resolveUpn(upn);
				if(StringUtils.isNotBlank(resolvedUpn) && (!resolvedUpn.equalsIgnoreCase(upn))){
					return findCalendarItemIdsInternal(resolvedUpn, startDate, endDate, calendarIds, newDepth);
				}else {
					//rethrow
					throw e0;
				}
			}catch(ExchangeExceededFindCountLimitRuntimeException e1) {
				log.warn("findCalendarItemIdsInternal(upn="+upn+",startDate="+startDate+",+endDate="+endDate+",...) ExceededFindCountLimit splitting request and trying again. - failure #"+newDepth);
				Set<ItemIdType> foundItems = new HashSet<ItemIdType>();
				List<Interval> intervals = ExchangeDateUtils.generateIntervals(startDate, endDate);
				for(Interval i: intervals) {
					foundItems.addAll(findCalendarItemIdsInternal(upn,i.getStart().toDate(), i.getEnd().toDate(),calendarIds,newDepth));
				}
				return foundItems;
			}catch(ExchangeWebServicesRuntimeException e2) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("findCalendarItemIdsInternal(upn="+upn+",startDate="+startDate+",+endDate="+endDate+",...) - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e2.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return findCalendarItemIdsInternal(upn, startDate, endDate, calendarIds, newDepth);
			}
		}
	}
	
	public Set<ItemIdType> findindFirstItemIdSet(String upn, Collection<FolderIdType> folderIds){
		FindItem request = getRequestFactory().constructFindFirstItemIdSet(folderIds);
		Pair<Set<ItemIdType>, Integer> pair =  findItemIdsInternal(upn, request, 0);
		return pair.getLeft();
	}
	
	
	public Set<ItemIdType> findItemIds(String upn, Collection<FolderIdType> folderIds){
		FindItem request = getRequestFactory().constructFindFirstItemIdSet(folderIds);
		Pair<Set<ItemIdType>, Integer> pair = findItemIdsInternal(upn, request, 0);
		return pair.getLeft();
	}
	
	public Set<ItemIdType> findAllItemIds(String upn, Collection<FolderIdType> folderIds){
		FindItem request = getRequestFactory().constructFindFirstItemIdSet(folderIds);
		Pair<Set<ItemIdType>, Integer> pair = findItemIdsInternal(upn, request, 0);
		Set<ItemIdType> itemIds = pair.getLeft();
		Integer nextOffset = pair.getRight();
		while(nextOffset > 0){
			request = getRequestFactory().constructFindNextItemIdSet(nextOffset, folderIds);
			pair = findItemIdsInternal(upn, request, 0);
			itemIds.addAll(pair.getLeft());
			nextOffset = pair.getRight();
		}
		return itemIds;
	}

	
	
	private Pair<Set<ItemIdType>, Integer> findItemIdsInternal(String upn, FindItem request, int depth){
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notNull(request, "request argument cannot be null");
		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("findCalendarItemIdsInternal(upn="+upn+",request="+request+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			try {
				FindItemResponse response = getWebServices().findItem(request);
				Pair<Set<ItemIdType>, Integer> parsed = getResponseUtils().parseFindItemIdResponse(response);

				return parsed;
			}catch(ExchangeInvalidUPNRuntimeException e0) {
				log.warn("findCalendarItemIdsInternal(upn="+upn+",request="+request+",...) ExchangeInvalidUPNRuntimeException.  Attempting to resolve valid upn... - failure #"+newDepth);
				
				String resolvedUpn = resolveUpn(upn);
				if(StringUtils.isNotBlank(resolvedUpn) && (!resolvedUpn.equalsIgnoreCase(upn))){
					return findItemIdsInternal(resolvedUpn, request, newDepth);
				}else {
					//rethrow
					throw e0;
				}
//			}catch(ExchangeExceededFindCountLimitRuntimeException e1) {
//				log.warn("findCalendarItemIdsInternal(upn="+upn+",request="+request+",...) ExceededFindCountLimit splitting request and trying again. - failure #"+newDepth);
//				Set<ItemIdType> foundItems = new HashSet<ItemIdType>();
//				List<Interval> intervals = DateHelp.generateIntervals(startDate, endDate);
//				for(Interval i: intervals) {
//					foundItems.addAll(findCalendarItemIdsInternal(upn,i.getStart().toDate(), i.getEnd().toDate(),calendarIds,newDepth));
//				}
//				return foundItems;
			}catch(ExchangeWebServicesRuntimeException e2) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("findCalendarItemIdsInternal(upn="+upn+",request="+request+",...) - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e2.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return findItemIdsInternal(upn, request, newDepth);
			}
		}
		
	}
	
	public Collection<CalendarItemType> getCalendarItems(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarFolderId){
		Set<ItemIdType> itemIds = findCalendarItemIds(upn, startDate, endDate, calendarFolderId);
		return getCalendarItems(upn, itemIds);
	}
	
	public Collection<CalendarItemType> getCalendarItems(String upn, Collection<ItemIdType> itemIds) {
		Set<CalendarItemType> calendarItems = new HashSet<CalendarItemType>();
		Set<ItemType> items = getItemsInternal(upn, itemIds, 0);
		for(ItemType item : items){
			if(item instanceof CalendarItemType){
				calendarItems.add( (CalendarItemType) item );
			}else{
				log.warn("non-calendarItemType will be excluded from result set");
			}
		}
		return calendarItems;
	}
	
	public CalendarItemType getCalendarItem(String upn, ItemIdType itemId) {
		Collection<CalendarItemType> items = getCalendarItems(upn, Collections.singleton(itemId));
		return DataAccessUtils.singleResult(items);
	}
	
	public Set<TaskType> getTaskItems(String upn, Set<ItemIdType> itemIds) {
		Set<TaskType> taskItems = new HashSet<TaskType>();
		Set<ItemType> items = getItemsInternal(upn, itemIds, 0);
		for(ItemType item : items){
			if(item instanceof TaskType){
				taskItems.add( (TaskType) item );
			}else{
				log.warn("non-TaskType will be excluded from result set");
			}
		}
		return taskItems;
	}
	
	private Set<ItemType> getItemsInternal(String upn, Collection<ItemIdType> itemIds, int depth){
		Set<ItemType> results = new HashSet<ItemType>();
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		if(CollectionUtils.isEmpty(itemIds)) {
			return results;
		}

		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("getItemsInternal(upn="+upn+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			GetItem request = getRequestFactory().constructGetItems(itemIds);	
			try {
				GetItemResponse response = getWebServices().getItem(request);
				return getResponseUtils().parseGetItemResponse(response);
			}catch(ExchangeWebServicesRuntimeException e) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("getItemsInternal - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return getItemsInternal(upn, itemIds, newDepth);
			}
		}
	}
	
	private ItemIdType createCalendarItemInternal(String upn, CalendarItemType calendarItem, int depth){
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notNull(calendarItem, "calendarItem argument cannot be empty");
	
		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("createCalendarItemInternal(upn="+upn+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			
			Set<CalendarItemType> singleton = Collections.singleton(calendarItem);
			CalendarItemCreateOrDeleteOperationType sendTo = CalendarItemCreateOrDeleteOperationType.SEND_TO_ALL_AND_SAVE_COPY;
			CreateItem request = getRequestFactory().constructCreateCalendarItem(singleton, sendTo, null);
			
			try {
				CreateItemResponse response = getWebServices().createItem(request);
				Set<ItemIdType> createdCalendarItems = getResponseUtils().parseCreateItemResponse(response);
				return DataAccessUtils.singleResult(createdCalendarItems);
			}catch(ExchangeWebServicesRuntimeException e) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("createCalendarItemInternal - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return createCalendarItemInternal(upn, calendarItem, newDepth);
			}
		}
		
	}

	public ItemIdType createCalendarItem(String upn, CalendarItemType calendarItem){
		return createCalendarItemInternal(upn, calendarItem, 0);
	}
	
	public FolderIdType createCalendarFolder(String upn, String displayName) {
		setContextCredentials(upn);
		log.debug("createCalendarFolder upn="+upn+", displayName="+displayName);
		//default implementation - null for extendedProperties
		CreateFolder createCalendarFolderRequest = getRequestFactory().constructCreateCalendarFolder(displayName, null);
		CreateFolderResponse createFolderResponse = getWebServices().createFolder(createCalendarFolderRequest);
		Set<FolderIdType> folders = getResponseUtils().parseCreateFolderResponse(createFolderResponse);
		return DataAccessUtils.singleResult(folders);
	}
	
	private boolean deleteCalendarItemsInternal(String upn, Collection<ItemIdType> itemIds, int depth) {
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notEmpty(itemIds, "itemIds argument cannot be empty");
	
		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("createCalendarItemInternal(upn="+upn+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			DeleteItem request = getRequestFactory().constructDeleteCalendarItems(itemIds, DisposalType.HARD_DELETE, CalendarItemCreateOrDeleteOperationType.SEND_TO_NONE);

			try {
				DeleteItemResponse response = getWebServices().deleteItem(request);
				boolean success = getResponseUtils().confirmSuccess(response);
				return success;
			
			}catch(ExchangeWebServicesRuntimeException e) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("deleteCalendarItemsInternal - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return deleteCalendarItemsInternal(upn, itemIds, newDepth);
			}
		}
	}
	
	public boolean deleteCalendarItems(String upn, Collection<ItemIdType> itemIds) {
		return deleteCalendarItemsInternal(upn, itemIds,0);
	}
	
	public Set<String> resolveEmailAddresses(String alias) {
		Validate.isTrue(StringUtils.isNotBlank(alias), "alias argument cannot be blank");
		setContextCredentials(getAdminUpn());
		ResolveNames request = getRequestFactory().constructResolveNames(alias);
		ResolveNamesResponse response = getWebServices().resolveNames(request);
		return getResponseUtils().parseResolveNamesResponse(response);
	}
	
	public String resolveUpn(String emailAddress) {
		Validate.isTrue(StringUtils.isNotBlank(emailAddress),"emailAddress argument cannot be blank");
		Validate.isTrue(EmailValidator.getInstance().isValid(emailAddress),"emailAddress argument must be valid");
		
		emailAddress = "smtp:"+emailAddress;
		
		Set<String> results = new HashSet<String>();
		Set<String> addresses = resolveEmailAddresses(emailAddress);
		for(String addr: addresses) {
			try {
				BaseFolderType primaryCalendarFolder = getPrimaryCalendarFolder(addr);
				if(null == primaryCalendarFolder) {
					throw new ExchangeRuntimeException("CALENDAR NOT FOUND");
				}else {
					results.add(addr);
				}
			}catch(RuntimeException e) {
				log.debug("resolveUpn -- "+addr+" NOT VALID. "+e.getMessage());
			}
		}
		if(CollectionUtils.isEmpty(results)) {
			throw new ExchangeRuntimeException("resolveUpn("+emailAddress+") failed -- no results.");
		}else {
			if(results.size() >1) {
				throw new ExchangeRuntimeException("resolveUpn("+emailAddress+") failed -- multiple results.");
			}else {
				return DataAccessUtils.singleResult(results);
			}
		}
	}	
    //================================================================================
    // ServerTimeZones
    //================================================================================
	public Set<TimeZoneDefinitionType> getServerTimeZones(String tzid, boolean fullTimeZoneData){
		GetServerTimeZones request = getRequestFactory().constructGetServerTimeZones(tzid, fullTimeZoneData);
		setContextCredentials(getAdminUpn());
		GetServerTimeZonesResponse response = getWebServices().getServerTimeZones(request);
		return getResponseUtils().parseGetServerTimeZonesResponse(response);
	}
	
	public Set<TimeZoneDefinitionType> getServerTimeZones(boolean fullTimeZoneData){
		return getServerTimeZones(null,fullTimeZoneData);
	}
	
	public TimeZoneDefinitionType getServerTimeZone(String tzid, boolean fullTimeZoneData){
		Set<TimeZoneDefinitionType> serverTimeZones = getServerTimeZones(tzid,fullTimeZoneData);
		return DataAccessUtils.singleResult(serverTimeZones);
	}	
	
	public boolean isEmpty(String upn, FolderIdType folderId){
		Set<ItemIdType> itemIds = findindFirstItemIdSet(upn, Collections.singleton(folderId));
		return CollectionUtils.isEmpty(itemIds);
	}
	
	/**
	 * The EmptyFolder operation empties folders in a mailbox. 
	 * Optionally, this operation enables you to delete the subfolders of the specified folder. 
	 * When a subfolder is deleted, the subfolder and the messages within the subfolder are deleted. 
	 * 
	 * *Note this method does not work for calendar or search folders: ERROR_CANNOT_EMPTY_FOLDER ... Emptying the calendar folder or search folder isn't permitted.
	 * 
	 * 
	 * @param upn
	 * @param folderId
	 * @return
	 */
	public boolean emptyFolder(String upn, boolean deleteSubFolders, DisposalType disposalType, BaseFolderIdType folderId){
		EmptyFolder request = getRequestFactory().constructEmptyFolder(deleteSubFolders, disposalType, Collections.singleton(folderId));
		setContextCredentials(upn);
		EmptyFolderResponse response = getWebServices().emptyFolder(request);
		return getResponseUtils().parseEmptyFolderResponse(response);
	}
	
	/**
	 * Deleting a calendarFolder with many (1k+) items is a problem.  You will always be throttled because the FindItemCount is 1000 and not configurable in Exchange Online.
	 * More info on throttling http://msdn.microsoft.com/en-us/library/office/jj945066(v=exchg.150).aspx
	 * 
	 * This method will never attempt to delete more than 500 items at once.
	 * 
	 * @param upn
	 * @param folderId
	 * @return
	 */
	public boolean emptyCalendarFolder(String upn, FolderIdType folderId){
		Integer deleteRequestCount =1;
		Set<ItemIdType> itemIds = findindFirstItemIdSet(upn, Collections.singleton(folderId));
		while(!itemIds.isEmpty()){
			List<ItemIdType> itemIdList = new ArrayList<ItemIdType>(itemIds);
			if(itemIdList.size() > 250){
				itemIdList = itemIdList.subList(0, 250);
			}
			StopWatch stopWatch = new StopWatch();
			stopWatch.start();
			log.info("emptyCalendarFolder(upn="+upn+") #"+deleteRequestCount+" deleting "+itemIdList.size()+ " calendar items");
			boolean result = deleteCalendarItems(upn, itemIdList);
			log.info("emptyCalendarFolder(upn="+upn+") #"+deleteRequestCount+" "+(result ? "Success" : "Failure")+" in "+stopWatch);
			itemIds = findindFirstItemIdSet(upn, Collections.singleton(folderId));
			deleteRequestCount++;
		}
		return true;
	}
	
	public boolean deleteCalendarFolder(String upn, FolderIdType folderId){
		boolean empty = emptyCalendarFolder(upn, folderId);
		if(empty){
			return deleteFolder(upn, DisposalType.SOFT_DELETE, folderId);					
		}
		return false;
	}
	
	/**
	 * First check the calendar for existing items, if the calendar folder is empty then delete it.
	 * @param upn
	 * @param folderId
	 * @return - True if the calender folder was deleted, false otherwise
	 */
	public boolean deleteEmptyCalendarFolder(String upn, FolderIdType folderId){
		return isEmpty(upn, folderId) ? deleteFolder(upn, DisposalType.SOFT_DELETE, folderId) : false;
	}
	
	/**
	 * Delete a calendar folder
	 * @param upn - the user 
	 * @param disposalType - how the deletion is performed
	 * @param folderId - the folder to delete
	 * @return
	 */
	public boolean deleteFolder(String upn, DisposalType disposalType, BaseFolderIdType folderId){
		DeleteFolder request = getRequestFactory().constructDeleteFolder(folderId, disposalType);
		setContextCredentials(upn);
		DeleteFolderResponse response = getWebServices().deleteFolder(request);
		return getResponseUtils().parseDeleteFolderResponse(response);
				
	}
	
	public ItemIdType createEmailMessage(List<String> recips, String replyTo, String subject, String messageBody, BodyTypeType bodyType, FolderIdType folderIdType){
		List<MessageType> messages = new ArrayList<MessageType>();
		MessageType messageType = new MessageType();
		
		//set one or more recipients to receive the message
		ArrayOfRecipientsType arrayOfRecips =new ArrayOfRecipientsType();
		for(String r: recips) {
			EmailAddressType emailAddressType =new EmailAddressType();
			emailAddressType.setEmailAddress(r);
			arrayOfRecips.getMailboxes().add(emailAddressType);
		}
		messageType.setToRecipients(arrayOfRecips);
		
		//set the replyTo address
		if(StringUtils.isNotBlank(replyTo)) {
			ArrayOfRecipientsType arrayOfReplyTos =new ArrayOfRecipientsType();
			SingleRecipientType fromType = new SingleRecipientType();
			EmailAddressType fromAddressType = new EmailAddressType();
			fromAddressType.setEmailAddress(replyTo);
			fromType.setMailbox(fromAddressType);
			arrayOfReplyTos.getMailboxes().add(fromAddressType);
			messageType.setReplyTo(arrayOfReplyTos);
		}
		
		//set the from address
		String from = getAdminSendAs();
		if(StringUtils.isNotBlank(from)){
			SingleRecipientType fromType = new SingleRecipientType();
			EmailAddressType fromAddressType = new EmailAddressType();
			fromAddressType.setEmailAddress(from);
			fromType.setMailbox(fromAddressType);
			messageType.setFrom(fromType);
		}
		//set the message body
		BodyType body = new BodyType();
		body.setBodyType(bodyType);
		body.setValue(messageBody);
		messageType.setBody(body);
		//set the subjec
		messageType.setSubject(subject);
		//message set as not read
		messageType.setIsRead(false);
		//add the message
		messages.add(messageType);
		
		//in this context we are impersonating an admin with send as rights for admin.sendas
		setContextCredentials(getAdminUpn());
		CreateItem request = getRequestFactory().constructCreateMessageItem(messages, folderIdType);
		CreateItemResponse response = getWebServices().createItem(request);
		Set<ItemIdType> items = getResponseUtils().parseCreateItemResponse(response);
		return DataAccessUtils.singleResult(items);
	}
	
	public boolean updateCalendarItemSetLegacyFreeBusy(String upn, CalendarItemType c){
		boolean itemUpdated = false;
		SetItemFieldType setField = getRequestFactory().constructSetCalendarItemLegacyFreeBusy(c);
		NonEmptyArrayOfItemChangesType changes = getRequestFactory().constructUpdateCalendarItemChanges(c, Collections.singleton(setField));
		UpdateItem request = getRequestFactory().constructUpdateCalendarItem(c, changes);
		UpdateItemResponse response = getWebServices().updateItem(request);
		Set<ItemIdType> itemIds = getResponseUtils().parseUpdateItemResponse(response);
		if(!CollectionUtils.isEmpty(itemIds)){
			ItemIdType itemId = DataAccessUtils.singleResult(itemIds);
			itemUpdated = c.getItemId().getId().equals(itemId.getId());
		}
		return itemUpdated;
	}
	
}
