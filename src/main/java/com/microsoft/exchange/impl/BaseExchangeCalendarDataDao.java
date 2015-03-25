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
import java.util.Random;
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
import org.springframework.dao.support.DataAccessUtils;
import org.springframework.util.CollectionUtils;

import com.microsoft.exchange.ExchangeDateUtils;
import com.microsoft.exchange.ExchangeRequestFactory;
import com.microsoft.exchange.ExchangeResponseUtils;
import com.microsoft.exchange.ExchangeWebServices;
import com.microsoft.exchange.exception.ExchangeExceededFindCountLimitRuntimeException;
import com.microsoft.exchange.exception.ExchangeInvalidUPNRuntimeException;
import com.microsoft.exchange.exception.ExchangeMissingEmailAddressRuntimeException;
import com.microsoft.exchange.exception.ExchangeRuntimeException;
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
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.ConnectingSIDType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.EmailAddressType;
import com.microsoft.exchange.types.ExtendedPropertyType;
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
 * Base class for interacting with {@link ExchangeWebServices}.
 * 
 * If you wish to get (or set) any {@link ExtendedPropertyType}s when retrieving
 * (or creating) items in Exchange you MUST provide your own implementation of {@link ExchangeRequestFactory}.
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
	private ExchangeRequestFactory requestFactory = new ExchangeRequestFactory();
	private ExchangeResponseUtils responseUtils = new ExchangeResponseUtilsImpl();
	private int maxRetries = 10;
	private static Random random = new Random();

	/**
	 * This is the principal/username which is used when performing non-user actions like:
	 *  {@link #getServerTimeZones(String, boolean)}, 
	 *  {@link #resolveEmailAddresses(String)} and
	 *  {@link #createEmailMessage(List, String, String, String, BodyTypeType, FolderIdType)}
	 */
	private String adminUpn;
	/**
	 * This property is used as the from address when sending mail.
	 * @see #createEmailMessage(List, String, String, String, BodyTypeType, FolderIdType)
	 */
	private String adminSendAs;
	
    //================================================================================
    // Getters/Setters 
    //================================================================================
	/**
	 * See: {@link #adminSendAs}
	 * @param adminSendAs the from address to use when sending email.
	 */
	public void setAdminSendAs(String adminSendAs){
		this.adminSendAs = adminSendAs;
	}
	
	/**
	 * See: {@link #adminSendAs}
	 * @return the from address used when sending email.
	 */
	public String getAdminSendAs(){
		return this.adminSendAs;
	}
	/**
	 * See: {@link #adminUpn}
	 * @param adminUpn the adminUpn to set
	 */
	public void setAdminUpn(String adminUpn){
		this.adminUpn = adminUpn;
	}
	/**
	 * See: {@link #adminUpn}
	 * @return the admin user principal name.
	 */
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
	@Autowired(required=false) @Qualifier("ewsClient")
	public void setWebServices(ExchangeWebServices exchangeWebServices) {
		this.webServices = exchangeWebServices;
	}

	/**
	 * @return the {@link ExchangeRequestFactory}
	 */
	public ExchangeRequestFactory getRequestFactory() {
		return requestFactory;
	}
	/**
	 * @param factory the requestFactory to set
	 */
	@Autowired(required=false)
	public void setRequestFactory(ExchangeRequestFactory exchangeRequestFactory) {
		this.requestFactory = exchangeRequestFactory;
	}
	
	/**
	 * @return the {@link ExchangeResponseUtils}
	 */
	public ExchangeResponseUtils getResponseUtils() {
		return responseUtils;
	}
	@Autowired(required=false)
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
	 * Function for generating an exponential backoff time. This method will
	 * never return less than 1000L. This method also adds a small random delay
	 * in an attempt to prevent threads from backing off/retrying in sync.
	 * 
	 * @param retryCount
	 *            the number of previously failed attempts.
	 * @return the number of ms to sleep for.
	 */
	public static long getWaitTimeExp(int retryCount) {
		Long baseMultiplier = 1000L;
		long waitTime = ((long) Math.pow(2, retryCount) * baseMultiplier);
		long rand =  ((long) random.nextInt(baseMultiplier.intValue())) + 1L;
	    return waitTime + rand;
	}
	
	/**
	 * @param upn
	 */
	protected void setContextCredentials(String upn) {
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		ConnectingSIDType connectingSID = new ConnectingSIDType();
		connectingSID.setPrincipalName(upn);
		ThreadLocalImpersonationConnectingSIDSourceImpl.setConnectingSID(connectingSID);
	}
	
	//================================================================================
    // GetFolder
    //================================================================================		
	/**
	 * Attempt to retrieve a {@link CalendarFolderType} from the Exchange server
	 * for the specified {@code upn} and {@link FolderIdType}
	 * 
	 * @param upn
	 * @param folderId
	 * @return the {@link CalendarFolderType} if found, otherwise <code>null</code>
	 */
	public CalendarFolderType getCalendarFolder(String upn, FolderIdType folderId) {
		BaseFolderType folder = getFolder(upn, folderId);
		CalendarFolderType calendarFolderType = null;
		if(folder instanceof CalendarFolderType) {
			calendarFolderType=  (CalendarFolderType) folder;
		}
		return calendarFolderType;
	}
	
	/**
	 * Attempt to retrieve a {@link TasksFolderType} from the Exchange server
	 * for the specified {@code upn} and {@link FolderIdType}
	 * 
	 * @param upn
	 * @param folderId
	 * @return the {@link TasksFolderType} if found, otherwise <code>null</code>
	 */
	public TasksFolderType getTaskFolder(String upn, FolderIdType folderId) {
		BaseFolderType folder = getFolder(upn, folderId);
		TasksFolderType taskFolderType = null;
		if(folder instanceof TasksFolderType) {
			taskFolderType =  (TasksFolderType) folder;
		}
		return taskFolderType;
	}
	
	/**
	 * Attempt to retrieve a {@link BaseFolderType} from the Exchange server
	 * for the specified {@code upn} and {@link FolderIdType}
	 * 
	 * @param upn
	 * @param folderId
	 * @return the {@link BaseFolderType} if found, otherwise <code>null</code>
	 */
	public BaseFolderType getFolder(String upn, FolderIdType folderIdType){
		setContextCredentials(upn);
		GetFolder getFolderRequest = getRequestFactory().constructGetFolderById(folderIdType);
		GetFolderResponse getFolderResponse = getWebServices().getFolder(getFolderRequest);
		Set<BaseFolderType> response = getResponseUtils().parseGetFolderResponse(getFolderResponse);
		return DataAccessUtils.singleResult(response);
	}
	
	/**
	 * Attempt to retrieve a {@link BaseFolderIdType} representing the primary item collection a {@link DistinguishedFolderIdNameType}
	 * 
	 * For instance, passing {@link DistinguishedFolderIdNameType#CALENDAR} to this function, the result should be the primary calendar folder.
	 *  
	 * @param upn
	 * @param parent
	 * @return
	 */
	protected BaseFolderType getPrimaryFolder(String upn, DistinguishedFolderIdNameType parent) {
		setContextCredentials(upn);
		GetFolder getFolderRequest = getRequestFactory().constructGetFolderByDistinguishedName(parent);
		GetFolderResponse getFolderResponse = getWebServices().getFolder(getFolderRequest);
		Set<BaseFolderType> response = getResponseUtils().parseGetFolderResponse(getFolderResponse);
		return DataAccessUtils.singleResult(response);
	}
	
	/**
	 * Obtain the primary calendar folder
	 * 
	 * @param upn
	 * @return {@link BaseFolderIdType}
	 */
	//TODO this should only return CalendarFolderType
	public BaseFolderType getPrimaryCalendarFolder(String upn){
		return getPrimaryFolder(upn, DistinguishedFolderIdNameType.CALENDAR);		
	}
	
	/**
	 * Obtain the primary task folder
	 * @param upn
	 * @return {@link BaseFolderIdType}
	 */
	//TODO this should only return TasksFolderType
	public BaseFolderType getPrimaryTaskFolder(String upn) {
		return getPrimaryFolder(upn, DistinguishedFolderIdNameType.TASKS);
	}
	
	/**
	 * Obtain the primary folder and all sub-folders for the given {@link DistinguishedFolderIdNameType}
	 * 
	 * @param upn
	 * @param parent
	 * @return a never null but possibly empty {@link Set} of {@link BaseFolderIdType}
	 */
	private Set<BaseFolderType> getFoldersByType(String upn, DistinguishedFolderIdNameType parent){
		Set<BaseFolderType> folders = new HashSet<BaseFolderType>();
		BaseFolderType baseFolderType = getPrimaryFolder(upn, parent);
		if(null != baseFolderType) {
			folders.add(baseFolderType);
		}
		Set<BaseFolderType> seondaryFolders = getSecondaryFolders(upn, parent);
		if(!CollectionUtils.isEmpty(seondaryFolders)) {
			for(BaseFolderType b: seondaryFolders ) {
				//TODO class comparison is not neccesary?
				if(baseFolderType.getClass().equals(b.getClass())) {
					folders.add(b);
				}
			}
		}
		return folders;
	}
	
	/**
	 * Obtain all sub-folders for given {@link DistinguishedFolderIdNameType}
	 * @param upn
	 * @param parent
	 * @return
	 */
	private Set<BaseFolderType> getSecondaryFolders(String upn, DistinguishedFolderIdNameType parent) {
		setContextCredentials(upn);
		FindFolder findFolderRequest = getRequestFactory().constructFindFolder(parent, DefaultShapeNamesType.ALL_PROPERTIES, FolderQueryTraversalType.DEEP,null);
		FindFolderResponse findFolderResponse = getWebServices().findFolder(findFolderRequest);
		return getResponseUtils().parseFindFolderResponse(findFolderResponse);
	}

	/**
	 * Obtain all {@link BaseFolderType}s representing calendar collections
	 * @param upn
	 * @return
	 */
	//TODO this method should only return Set<CalendarItemType>
	public Set<BaseFolderType> getAllCalendarFolders(String upn) {
		Validate.notEmpty(upn, "upn cannnot be empty");
		return getFoldersByType(upn, DistinguishedFolderIdNameType.CALENDAR);
	}
	
	/**
	 * Obtain all {@link BaseFolderType}s representing Task collections
	 * @param upn
	 * @return
	 */
	//TODO this method should only return Set<TasksItemType>
	public Set<BaseFolderType> getAllTaskFolders(String upn) {
		Validate.notEmpty(upn, "upn cannnot be empty");
		return getFoldersByType(upn, DistinguishedFolderIdNameType.TASKS);
	}
	
	/**
	 * Obtain the {@link FolderIdType} for a {@link CalendarFolderType} given the display name of the calendar.
	 * @param upn
	 * @param calendarName
	 * @return
	 */
	//TODO this should definately not throw a RuntimeException
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

	/**
	 * Obtain all {@link CalendarFolderType}s from the exchange server and return a map
	 * where the key represents the FolderId and the value represents the
	 * {@link CalendarFolderType}s dispaly name
	 * 
	 * @param upn
	 * @return
	 */
	public Map<String, String> getCalendarFolderMap(String upn){
		Map<String, String> calendarsMap = new HashMap<String, String>();
		Set<BaseFolderType> allCalendarFolders = getAllCalendarFolders(upn);
		log.debug("getCalendarFolderMap found "+allCalendarFolders.size());
		Integer i= 1;
		for(BaseFolderType folderType: allCalendarFolders) {
			String name = folderType.getDisplayName();
			String id = folderType.getFolderId().getId();
			log.debug("CalendarFolderMap "+upn+" "+i+": "+name);
			calendarsMap.put(id, name);
			i++;
		}
		return calendarsMap;
	}
	
	/**
	 * Obtain all {@link TasksFolderType}s from the exchange server and return a map
	 * where the key represents the FolderId and the value represents the
	 * {@link TasksFolderType}s dispaly name
	 * 
	 * @param upn
	 * @return
	 */
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
	
	//================================================================================
    // FindItem
    //================================================================================	
	
	/**
	 * Find all {@link ItemIdType} within the primary {@link CalendarFolderType} between {@code startDate} and {@code endDate}
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	public Set<ItemIdType> findCalendarItemIds(String upn, Date startDate, Date endDate){
		return findCalendarItemIdsInternal(upn, startDate, endDate, null, 0);
	}
	
	/**
	 * Find all {@link ItemIdType} within the specified {@link FolderIdType}s between {@code startDate} and {@code endDate}
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	public Set<ItemIdType> findCalendarItemIds(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarIds) {
		return findCalendarItemIdsInternal(upn, startDate, endDate, calendarIds, 0);
	}
	
	/**
	 * This method issues a {@link FindItem} request using a {@link CalendarViewType} to obtain identifiers for all {@link CalendarItemType}s between {@code startDate} and {@code endDate}
	 * 
	 * Note: CalendarView element returns single calendar items and all occurrences.  In other words, this method expands recurrence for you.
	 * 
	 * @see <a href='http://msdn.microsoft.com/en-us/library/office/aa566107(v=exchg.140).aspx'>FindItem Operation</a>
	 * 
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @param calendarIds - if omitted the primary calendar folder will be targeted
	 * @param depth
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	private Set<ItemIdType> findCalendarItemIdsInternal(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarIds, int depth){
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notNull(startDate, "startDate argument cannot be null");
		Validate.notNull(endDate, "endDate argument cannot be null");
		//if folderIds is empty the primary calendar folder will be targeted
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
			}catch(ExchangeRuntimeException e2) {
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
	
	/**
	 * Locate the {@link ItemIdType}s within the specified folder(s).
	 * This method will return a maximum of {@link BaseExchangeCalendarDataDao#getRequestFactory()}.getMaxFindItems()
	 * 
	 * @param upn
	 * @param folderIds
	 * @return {@link Set} of {@link ItemIdType}
	 */
	public Set<ItemIdType> findItemIds(String upn, Collection<FolderIdType> folderIds){
		FindItem request = getRequestFactory().constructFindFirstItemIdSet(folderIds);
		Pair<Set<ItemIdType>, Integer> pair = findItemIdsInternal(upn, request, 0);
		return pair.getLeft();
	}
	
	/**
	 * Obtain all {@link ItemIdType}s within a specified {@link FolderIdType} by
	 * repeatedly callling {@link FindItem} and paging the results
	 * 
	 * @param upn
	 * @param folderIds
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
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

	/**
	 * 
	 * @param upn
	 * @param request
	 * @param depth
	 * @return
	 */
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
			}catch(ExchangeRuntimeException e2) {
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
	
	//================================================================================
    // GetItem
    //================================================================================	
	
	/**
	 * Obtain all {@link CalendarItemType}s found within the specified
	 * {@code calendarFolderId} which intersect the date range from
	 * {@code startDate} to {@code endDate}
	 * 
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @param calendarFolderId
	 * @return a never null but possibly empty {@link Collection} of {@link CalendarItemType}
	 */
	public Collection<CalendarItemType> getCalendarItems(String upn, Date startDate, Date endDate, Collection<FolderIdType> calendarFolderId){
		Set<ItemIdType> itemIds = findCalendarItemIds(upn, startDate, endDate, calendarFolderId);
		return getCalendarItems(upn, itemIds);
	}
	
	/**
	 * Obtain a {@link CalendarItemType} for each {@link ItemIdType} specified.
	 * @param upn
	 * @param itemIds
	 * @return a never null but possibly empty {@link Collection} of {@link CalendarItemType}
	 */
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
	
	/**
	 * Get the {@link CalendarItemType} for the specified {@link ItemIdType} 
	 * @param upn
	 * @param itemId
	 * @return a {@link CalendarItemType} if found, <code>null</code> otherwise
	 */
	public CalendarItemType getCalendarItem(String upn, ItemIdType itemId) {
		Collection<CalendarItemType> items = getCalendarItems(upn, Collections.singleton(itemId));
		return DataAccessUtils.singleResult(items);
	}
	
	/**
	 * Obtain a {@link TaskType} for each {@link ItemIdType} specified.
	 * @param upn
	 * @param itemIds
	 * @return a never null but possibly empty {@link Collection} of {@link TaskType}
	 */
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
	
	/**
	 * 
	 * @param upn the UserPrincipalName
	 * @param itemIds the {@link ItemIdType}s
	 * @param depth the number of recursive calls.
	 * @return a never null but possbly empty {@link Set} of {@link ItemType}s
	 */
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
			}catch(ExchangeRuntimeException e) {
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
	
	//================================================================================
    // CreateItem
    //================================================================================	
	
	/**
	 * Create the {@link CalendarItemType} on the exchange server
	 * @param upn
	 * @param calendarItem
	 * @param depth
	 * @return {@link ItemIdType}
	 */
	private ItemIdType createCalendarItemInternal(String upn, CalendarItemType calendarItem, FolderIdType calendarFolderId, int depth){
		Validate.notNull(calendarItem, "calendarItem argument cannot be empty");
		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("createCalendarItemInternal(upn="+upn+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			Set<CalendarItemType> singleton = Collections.singleton(calendarItem);
			CreateItem request = getRequestFactory().constructCreateCalendarItem(singleton, calendarFolderId);
			try {
				CreateItemResponse response = getWebServices().createItem(request);
				Set<ItemIdType> createdCalendarItems = getResponseUtils().parseCreateItemResponse(response);
				return DataAccessUtils.singleResult(createdCalendarItems);
			}catch(ExchangeRuntimeException e) {
				long backoff = getWaitTimeExp(newDepth);
				log.warn("createCalendarItemInternal - failure #"+newDepth+". Sleeping for "+backoff+" before retry. " +e.getMessage());
				try {
					Thread.sleep(backoff);
				} catch (InterruptedException e1) {
					log.warn("InterruptedException="+e1);
				}
				return createCalendarItemInternal(upn, calendarItem, calendarFolderId, newDepth);
			}
		}
		
	}

	/**
	 * Create the {@link CalendarItemType} on the exchange server
	 * @param upn
	 * @param calendarItem
	 * @return {@link ItemIdType}
	 */
	public ItemIdType createCalendarItem(String upn, CalendarItemType calendarItem){
		return createCalendarItem(upn, calendarItem, null);
	}
	
	public ItemIdType createCalendarItem(String upn, CalendarItemType calendarItem, FolderIdType calendarFolderId){
		return createCalendarItemInternal(upn, calendarItem, calendarFolderId, 0);
	}
	
	/**
	 * Create a {@link CalendarFolderType} with the specified {@code displayName}
	 * @param upn
	 * @param displayName
	 * @return {@link FolderIdType}
	 */
	public FolderIdType createCalendarFolder(String upn, String displayName) {
		setContextCredentials(upn);
		log.debug("createCalendarFolder upn="+upn+", displayName="+displayName);
		CreateFolder createCalendarFolderRequest = getRequestFactory().constructCreateCalendarFolder(displayName, null);
		CreateFolderResponse createFolderResponse = getWebServices().createFolder(createCalendarFolderRequest);
		Set<FolderIdType> folders = getResponseUtils().parseCreateFolderResponse(createFolderResponse);
		return DataAccessUtils.singleResult(folders);
	}
	
	/**
	 * Create and send an Email (i.e. {@link MessageType}) 
	 * @param recips - addresses to send the message to
	 * @param replyTo - a single email address, to whom the reply will be addressed
	 * @param subject - the email messages subject
	 * @param messageBody
	 * @param bodyType 
	 * @param folderIdType
	 * @return {@link ItemIdType}
	 */
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
	
	//================================================================================
    // DeleteItem
    //================================================================================	
	private boolean deleteCalendarItemsInternal(String upn, Collection<ItemIdType> itemIds, int depth) {
		Validate.isTrue(StringUtils.isNotBlank(upn), "upn argument cannot be blank");
		Validate.notEmpty(itemIds, "itemIds argument cannot be empty");
	
		int newDepth = depth +1;
		if(depth > getMaxRetries()) {
			throw new ExchangeRuntimeException("deleteCalendarItemsInternal(upn="+upn+",...) failed "+getMaxRetries()+ " consecutive attempts.");
		}else {
			setContextCredentials(upn);
			DeleteItem request = getRequestFactory().constructDeleteCalendarItems(itemIds);
			try {
				DeleteItemResponse response = getWebServices().deleteItem(request);
				boolean success = getResponseUtils().confirmSuccess(response);
				return success;
			
			}catch(ExchangeRuntimeException e) {
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
	
	/**
	 * Delete {@link CalendarItemType}s for the user.
	 * @param upn - the userPrincipalName identifies the user to delete {@link CalendarItemType}s for.
	 * @param itemIds - the {@link ItemIdType}s for the {@link CalendarItemType}s to delete.
	 * 
	 * @see ExchangeRequestFactory#constructDeleteCalendarItems(Collection, DisposalType, CalendarItemCreateOrDeleteOperationType)
	 * @return
	 */
	public boolean deleteCalendarItems(String upn, Collection<ItemIdType> itemIds) {
		return deleteCalendarItemsInternal(upn, itemIds,0);
	}

	//================================================================================
    // ResolveNames
    //================================================================================	

	/**
	 * Search the exchange server for any contacts with a name similar to
	 * {@code alias} and return all corresponding SMTP addresses
	 * 
	 * @param alias
	 * @return
	 */
	public Set<String> resolveEmailAddresses(String alias) {
		Set<String> smtpAddresses = Collections.emptySet();
		Validate.isTrue(StringUtils.isNotBlank(alias), "alias argument cannot be blank");
		setContextCredentials(getAdminUpn());
		ResolveNames request = getRequestFactory().constructResolveNames(alias);
		ResolveNamesResponse response = getWebServices().resolveNames(request);
		try{
			smtpAddresses = getResponseUtils().parseResolveNamesResponse(response);
		}catch(ExchangeMissingEmailAddressRuntimeException e){
			request = getRequestFactory().constructResolveNamesWithDistinguishedFolderId(alias);
			response = getWebServices().resolveNames(request);
			smtpAddresses = getResponseUtils().parseResolveNamesResponse(response);
		}
		return smtpAddresses;
	}
	
	/**
	 * Search the exchange server for any contacts with an email address that matches
	 * {@code emailAddress} and return only the UserPrincipalName
	 * 
	 * @param alias
	 * @return a {@link String} representing the UPN
	 */
	public String resolveUpn(String emailAddress) {
		Validate.isTrue(StringUtils.isNotBlank(emailAddress),"emailAddress argument cannot be blank");
		Validate.isTrue(EmailValidator.getInstance().isValid(emailAddress),"emailAddress argument must be valid");
		
		emailAddress = ExchangeRequestFactory.SMTP + emailAddress;
		Map<BaseFolderType, String> resultMap = new HashMap<BaseFolderType, String>();
		Set<String> addresses = resolveEmailAddresses(emailAddress);
		for(String addr: addresses) {
			try {
				BaseFolderType primaryCalendarFolder = getPrimaryCalendarFolder(addr);
				if(null == primaryCalendarFolder) {
					throw new ExchangeRuntimeException("CALENDAR NOT FOUND");
				}else {
					resultMap.put(primaryCalendarFolder, addr);
				}
			}catch(RuntimeException e) {
				log.debug("resolveUpn -- "+addr+" NOT VALID. "+e.getMessage());
			}
		}
		if(CollectionUtils.isEmpty(resultMap)) {
			throw new ExchangeRuntimeException("resolveUpn("+emailAddress+") failed -- no results.");
		}else {
			if(resultMap.isEmpty()) {
				throw new ExchangeRuntimeException("resolveUpn("+emailAddress+") failed -- multiple results.");
			}else {
				//just return the first entry.
				BaseFolderType key = resultMap.keySet().iterator().next();
				return resultMap.get(key);
			}
		}
	}	
    //================================================================================
    // ServerTimeZones
    //================================================================================
	/**
	 * Get {@link TimeZoneDefinitionType}s from the Exchange Server
	 * @param tzid - if specified the server will only return a single matching {@link TimeZoneDefinitionType}
	 * @param fullTimeZoneData -
	 * @return a never null but possibly empty {@link Set} of {@link TimeZoneDefinitionType}
	 */
	public Set<TimeZoneDefinitionType> getServerTimeZones(String tzid, boolean fullTimeZoneData){
		GetServerTimeZones request = getRequestFactory().constructGetServerTimeZones(tzid, fullTimeZoneData);
		setContextCredentials(getAdminUpn());
		GetServerTimeZonesResponse response = getWebServices().getServerTimeZones(request);
		return getResponseUtils().parseGetServerTimeZonesResponse(response);
	}
	
	/**
	 * Get all available {@link TimeZoneDefinitionType}s from the Exchange Server
	 * @param fullTimeZoneData
	 * @return a never null but possibly empty {@link Set} of {@link TimeZoneDefinitionType}
	 */
	public Set<TimeZoneDefinitionType> getServerTimeZones(boolean fullTimeZoneData){
		return getServerTimeZones(null,fullTimeZoneData);
	}
	
	/**
	 * Get the {@link TimeZoneDefinitionType} with a timeZoneId of {@code tzid}
	 * @param tzid
	 * @param fullTimeZoneData
	 * @return {@link TimeZoneDefinitionType} if found, <code>null</code> otherwise
	 */
	public TimeZoneDefinitionType getServerTimeZone(String tzid, boolean fullTimeZoneData){
		Set<TimeZoneDefinitionType> serverTimeZones = getServerTimeZones(tzid,fullTimeZoneData);
		return DataAccessUtils.singleResult(serverTimeZones);
	}	
	
	//================================================================================
    // EmptyFolder
    //================================================================================	
	/**
	 * Attempt to find items within the specified {@code folderId}.
	 * @param upn
	 * @param folderId
	 * @return true only if the folder is completely empty
	 */
	public boolean isEmpty(String upn, FolderIdType folderId){
		Set<ItemIdType> itemIds = findItemIds(upn, Collections.singleton(folderId));
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
	public boolean emptyFolder(String upn, boolean deleteSubFolders,  BaseFolderIdType folderId){
		EmptyFolder request = getRequestFactory().constructEmptyFolder(deleteSubFolders,  Collections.singleton(folderId));
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
		Set<ItemIdType> itemIds = findItemIds(upn, Collections.singleton(folderId));
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
			itemIds = findItemIds(upn, Collections.singleton(folderId));
			deleteRequestCount++;
		}
		return true;
	}
	
	public boolean purgeCancelledCalendarItems(String upn, FolderIdType folderId){
		FindItem request = getRequestFactory().constructIndexedPageViewFindItemCancelledCalendarItemIds(Collections.singleton(folderId));
		FindItemResponse response = getWebServices().findItem(request);
		Pair<Set<ItemIdType>, Integer> results = getResponseUtils().parseFindItemIdResponse(response);
		Set<ItemIdType> itemIds = results.getLeft();
		Integer nextOffset = results.getRight();
		while(itemIds.size() > 0){
			if(deleteCalendarItems(upn, itemIds)){
				if(nextOffset > 0){
					request = getRequestFactory().constructIndexedPageViewFindItemCancelledCalendarItemIds(nextOffset, Collections.singleton(folderId));
					response = getWebServices().findItem(request);
					results = getResponseUtils().parseFindItemIdResponse(response);
				
					itemIds = results.getLeft();
					nextOffset = results.getRight();
				}
			}else{
				return false;
			}
		}
		return true;
	}
	//================================================================================
    // DeleteFolder
    //================================================================================	
	/**
	 * Delete the specified {@code folderId} and all items contained wihtin it.
	 * 
	 * @param upn
	 * @param folderId
	 * @return
	 */
	public boolean deleteCalendarFolder(String upn, FolderIdType folderId){
		boolean empty = emptyCalendarFolder(upn, folderId);
		if(empty){
			return deleteFolder(upn, folderId);					
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
		return isEmpty(upn, folderId) ? deleteFolder(upn, folderId) : false;
	}
	
	/**
	 * Delete a calendar folder
	 * @param upn - the user 
	 * @param disposalType - how the deletion is performed
	 * @param folderId - the folder to delete
	 * @return
	 */
	public boolean deleteFolder(String upn, BaseFolderIdType folderId){
		DeleteFolder request = getRequestFactory().constructDeleteFolder(folderId);
		setContextCredentials(upn);
		DeleteFolderResponse response = getWebServices().deleteFolder(request);
		return getResponseUtils().parseDeleteFolderResponse(response);
	}
	
	/**
	 * Force the exchange server to update the {@link CalendarItemType}s subject field
  	 * @param upn
	 * @param c
	 * @return- true if the update succeded
	 */
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
