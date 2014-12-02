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
package com.microsoft.exchange;

import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.Validate;
import org.springframework.util.CollectionUtils;

import com.microsoft.exchange.messages.CreateFolder;
import com.microsoft.exchange.messages.CreateItem;
import com.microsoft.exchange.messages.DeleteFolder;
import com.microsoft.exchange.messages.DeleteItem;
import com.microsoft.exchange.messages.EmptyFolder;
import com.microsoft.exchange.messages.FindFolder;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.messages.GetFolder;
import com.microsoft.exchange.messages.GetItem;
import com.microsoft.exchange.messages.GetServerTimeZones;
import com.microsoft.exchange.messages.GetUserAvailabilityRequest;
import com.microsoft.exchange.messages.GetUserConfiguration;
import com.microsoft.exchange.messages.ResolveNames;
import com.microsoft.exchange.messages.UpdateFolder;
import com.microsoft.exchange.types.AcceptItemType;
import com.microsoft.exchange.types.AffectedTaskOccurrencesType;
import com.microsoft.exchange.types.ArrayOfCalendarPermissionsType;
import com.microsoft.exchange.types.ArrayOfMailboxData;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.BaseItemIdType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemCreateOrDeleteOperationType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.CalendarPermissionLevelType;
import com.microsoft.exchange.types.CalendarPermissionReadAccessType;
import com.microsoft.exchange.types.CalendarPermissionSetType;
import com.microsoft.exchange.types.CalendarPermissionType;
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.DeclineItemType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.DistinguishedFolderIdType;
import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.FolderQueryTraversalType;
import com.microsoft.exchange.types.FolderType;
import com.microsoft.exchange.types.FreeBusyViewOptions;
import com.microsoft.exchange.types.IndexedPageViewType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.MailboxData;
import com.microsoft.exchange.types.MessageDispositionType;
import com.microsoft.exchange.types.MessageType;
import com.microsoft.exchange.types.NonEmptyArrayOfAllItemsType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseFolderIdsType;
import com.microsoft.exchange.types.NonEmptyArrayOfFoldersType;
import com.microsoft.exchange.types.NonEmptyArrayOfTimeZoneIdType;
import com.microsoft.exchange.types.ObjectFactory;
import com.microsoft.exchange.types.PathToExtendedFieldType;
import com.microsoft.exchange.types.PathToUnindexedFieldType;
import com.microsoft.exchange.types.PermissionActionType;
import com.microsoft.exchange.types.ResolveNamesSearchScopeType;
import com.microsoft.exchange.types.ResponseTypeType;
import com.microsoft.exchange.types.RestrictionType;
import com.microsoft.exchange.types.SearchFolderTraversalType;
import com.microsoft.exchange.types.SearchFolderType;
import com.microsoft.exchange.types.SearchParametersType;
import com.microsoft.exchange.types.SuggestionsViewOptions;
import com.microsoft.exchange.types.TaskType;
import com.microsoft.exchange.types.TasksFolderType;
import com.microsoft.exchange.types.TentativelyAcceptItemType;
import com.microsoft.exchange.types.TimeZone;
import com.microsoft.exchange.types.UnindexedFieldURIType;
import com.microsoft.exchange.types.UserConfigurationNameType;
import com.microsoft.exchange.types.UserIdType;
import com.microsoft.exchange.types.WellKnownResponseObjectType;

public class ExchangeRequestFactory extends BaseExchangeRequestFactory{
	
	Collection<PathToExtendedFieldType> itemExtendedPropertyPaths = new HashSet<PathToExtendedFieldType>();
	Collection<PathToExtendedFieldType> folderExtendedPropertyPaths = new HashSet<PathToExtendedFieldType>();
	
	/* (non-Javadoc)
	 * @see com.microsoft.exchange.BaseExchangeRequestFactory#getItemExtendedPropertyPaths()
	 */
	@Override
	public Collection<PathToExtendedFieldType> getItemExtendedPropertyPaths() {
		return itemExtendedPropertyPaths;
	}
	
	public void setItemExtendedPropertyPaths(Collection<PathToExtendedFieldType> itemPaths) {
		itemExtendedPropertyPaths = itemPaths;
	}

	/* (non-Javadoc)
	 * @see com.microsoft.exchange.BaseExchangeRequestFactory#getFolderExtendedPropertyPaths()
	 */
	@Override
	public Collection<PathToExtendedFieldType> getFolderExtendedPropertyPaths() {
		return folderExtendedPropertyPaths;
	}	
	
	public void setFolderExtendedPropertyPaths(Collection<PathToExtendedFieldType> folderPaths) {
		this.folderExtendedPropertyPaths = folderPaths;
	}
	
	/**
	 * Construct a {@link ResolveNames} request which will search the Exchange
	 * server using the
	 * {@link ResolveNamesSearchScopeType#ACTIVE_DIRECTORY_CONTACTS} scope. This
	 * method uses {@link DefaultShapeNamesType#ALL_PROPERTIES} to return all
	 * properties for the contact matching the alias
	 * 
	 * @param alias
	 * @return {@link ResolveNames}
	 */
	public ResolveNames constructResolveNames(String alias) {
		return constructResolveNames(alias, true, ResolveNamesSearchScopeType.ACTIVE_DIRECTORY_CONTACTS, DefaultShapeNamesType.ALL_PROPERTIES);
	}
	
	/**
	 * Construct a {@link GetServerTimeZones} request
	 * @param tzid - optional time zone id.  if specified the response will only contain information regarding this tz
	 * @param returnFullTimeZoneData 
	 * @return {@link GetServerTimeZones}
	 */
	public GetServerTimeZones constructGetServerTimeZones(String tzid, boolean returnFullTimeZoneData){
		GetServerTimeZones request = new GetServerTimeZones();
		if(StringUtils.isNotBlank(tzid)){
			NonEmptyArrayOfTimeZoneIdType tzidArray = new NonEmptyArrayOfTimeZoneIdType();
			tzidArray.getIds().add(tzid);
			request.setIds(tzidArray);
		}
		request.setReturnFullTimeZoneData(returnFullTimeZoneData);
		return request;
	}
	
	/**
	 * Construct {@link GetUserConfiguration} request
	 * @param name
	 * @param distinguishedFolderIdType
	 * @return {@link GetUserConfiguration}
	 */
	public GetUserConfiguration constructGetUserConfiguration(String name, DistinguishedFolderIdType distinguishedFolderIdType) {
		GetUserConfiguration getUserConfiguration = new GetUserConfiguration();
		UserConfigurationNameType userConfigurationNameType = new UserConfigurationNameType();
		userConfigurationNameType.setDistinguishedFolderId(distinguishedFolderIdType);
		userConfigurationNameType.setName(name);
		getUserConfiguration.setUserConfigurationName(userConfigurationNameType);
		return getUserConfiguration;
	}

	/**
	 * Constructs a {@link GetUserAvailabilityRequest} object
	 * @param mailboxData
	 * @param freeBusyView
	 * @param suggestionsView
	 * @param timeZone
	 * @return {@link GetUserAvailabilityRequest}
	 */
	public GetUserAvailabilityRequest constructGetUserAvailabilityRequest(Collection<? extends MailboxData> mailboxData, FreeBusyViewOptions freeBusyView, SuggestionsViewOptions suggestionsView, TimeZone timeZone){
		GetUserAvailabilityRequest request = new GetUserAvailabilityRequest();
		if(!CollectionUtils.isEmpty(mailboxData)){
			ArrayOfMailboxData arrayOfMailboxData = new ArrayOfMailboxData();
			arrayOfMailboxData.getMailboxDatas().addAll(mailboxData);
			request.setMailboxDataArray(arrayOfMailboxData);
		}
		if(null != suggestionsView){
			request.setSuggestionsViewOptions(suggestionsView);
		}
		if(null != freeBusyView){
			request.setFreeBusyViewOptions(freeBusyView);
		}
		if(null != timeZone){
			request.setTimeZone(timeZone);
		}
		return request;
	}

    //================================================================================
    // CreateFolder
    //================================================================================
	/**
	 * Constructs a {@link CreateFolder} request specific for createing a {@link SearchFolderType}
	 * 
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/dn592094(v=exchg.150).aspx">How to: Work with search folders by using EWS in Exchange</a>
	 * @category Create Search Folder
	 * 
	 * @param displayName
	 * @param searchRoot
	 * @param restriction
	 * @return {@link CreateFolder}
	 */
	public CreateFolder constructCreateSearchFolder(String displayName,
			DistinguishedFolderIdNameType searchRoot,
			RestrictionType restriction) {
		CreateFolder createFolder = new CreateFolder();

		// create new searchFolderType
		SearchFolderType searchFolderType = new SearchFolderType();

		// create search parameters
		SearchParametersType searchParameters = new SearchParametersType();

		// search folders recursively
		searchParameters.setTraversal(SearchFolderTraversalType.DEEP);

		NonEmptyArrayOfBaseFolderIdsType baseFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
		baseFolderIds.getFolderIdsAndDistinguishedFolderIds().add(
				getParentDistinguishedFolderId(searchRoot));

		// set the baase of the search
		searchParameters.setBaseFolderIds(baseFolderIds);

		// set the search restriction
		searchParameters.setRestriction(restriction);

		// add search parameters to folder
		searchFolderType.setSearchParameters(searchParameters);

		// set the search folder display name
		searchFolderType.setDisplayName(displayName);

		// add searchFolder to CreatFolder request
		NonEmptyArrayOfFoldersType nonEmptyArrayOfFoldersType = new NonEmptyArrayOfFoldersType();
		nonEmptyArrayOfFoldersType
				.getFoldersAndCalendarFoldersAndContactsFolders().add(
						searchFolderType);
		createFolder.setFolders(nonEmptyArrayOfFoldersType);

		createFolder
				.setParentFolderId(getParentTargetFolderId(DistinguishedFolderIdNameType.SEARCHFOLDERS));

		return createFolder;
	}
	
	/**
	 * Construct a {@link CreateFolder} request specific for creating a {@link CalendarFolderType}
	 * @category Create Calendar Folder
	 * 
	 * @param displayName
	 * @param exProps - {@link ExtendedPropertyType}s to set on the created {@link CalendarFolderType}
	 * @return {@link CreateFolder}
	 */
	public CreateFolder constructCreateCalendarFolder(String displayName, Collection<ExtendedPropertyType> exProps) {
		Validate.isTrue(StringUtils.isNotBlank(displayName),"displayName argument cannot be empty");
		CalendarFolderType calendarFolder = new CalendarFolderType();
		calendarFolder.setDisplayName(displayName);
		if(!CollectionUtils.isEmpty(exProps)){
			calendarFolder.getExtendedProperties().addAll(exProps);
		}
		return constructCreateFolder(DistinguishedFolderIdNameType.CALENDAR, Collections.singleton(calendarFolder));
	}

	/**
	 * Constructs a {@link CreateFolder} request specific for creating {@link TasksFolderType}s.
	 * @category Create Task Folder
	 * @param displayName
	 * @param exProps  {@link ExtendedPropertyType}s to set on the created {@link TasksFolderType}
	 * @return {@link CreateFolder}
	 */
	public CreateFolder constructCreateTaskFolder(String displayName,
			Collection<ExtendedPropertyType> exProps) {
		Validate.isTrue(StringUtils.isNotBlank(displayName),"displayName argument cannot be empty");
		TasksFolderType taskFolder = new TasksFolderType();
		taskFolder.setDisplayName(displayName);
		if(!CollectionUtils.isEmpty(exProps)){
			taskFolder.getExtendedProperties().addAll(exProps);
		}
		return constructCreateFolder(DistinguishedFolderIdNameType.TASKS, Collections.singleton(taskFolder));
	}

	/**
	 * Attempt to create a calendar group i.e. a folder that may contain a number of sub calendars.
	 * I don't think you can create a calendar group using EWS, though it's possible this code is just wrong.
	 * @param upn
	 * @param displayName
	 * @return {@link CreateFolder}
	 */
	@Deprecated
	public CreateFolder constructCreateCalendarFolderGroup(String upn,
			String displayName) {
		CalendarFolderType calendarFolderType = new CalendarFolderType();
		calendarFolderType.setDisplayName(displayName);
		CalendarPermissionType calendarPermissionType = new CalendarPermissionType();
		UserIdType userId = new UserIdType();
		userId.setPrimarySmtpAddress(upn);

		calendarPermissionType.setUserId(userId);
		calendarPermissionType.setCanCreateSubFolders(true);
		calendarPermissionType.setIsFolderOwner(true);
		calendarPermissionType.setIsFolderContact(true);
		calendarPermissionType.setIsFolderVisible(true);
		calendarPermissionType.setEditItems(PermissionActionType.ALL);
		calendarPermissionType.setDeleteItems(PermissionActionType.ALL);
		calendarPermissionType
				.setReadItems(CalendarPermissionReadAccessType.FULL_DETAILS);
		calendarPermissionType
				.setCalendarPermissionLevel(CalendarPermissionLevelType.OWNER);

		ArrayOfCalendarPermissionsType arrayOfCalendarPermissionsType = new ArrayOfCalendarPermissionsType();
		arrayOfCalendarPermissionsType.getCalendarPermissions().add(
				calendarPermissionType);

		CalendarPermissionSetType calendarPermissionSetType = new CalendarPermissionSetType();
		calendarPermissionSetType
				.setCalendarPermissions(arrayOfCalendarPermissionsType);

		calendarFolderType.setPermissionSet(calendarPermissionSetType);
		return constructCreateFolder(DistinguishedFolderIdNameType.CALENDAR,
				Collections.singleton(calendarFolderType));
	}

	//================================================================================
    // GetFolder
    //================================================================================	
	
	/**
	 * Construct a {@link GetFolder} request for the given {@link DistinguishedFolderIdNameType}
	 * @category Get Distinguished Folder
	 * @param name
	 * @return{@link GetFolder}
	 */
	public GetFolder constructGetFolderByDistinguishedName(DistinguishedFolderIdNameType name) {
		DistinguishedFolderIdType parentDistinguishedFolderId = getParentDistinguishedFolderId(name);
		return constructGetFolderById(parentDistinguishedFolderId);
	}

	/**
	 * Construct a {@link GetFolder} request for a given {@link BaseFolderIdType}
	 * @category Get Base Folder
	 * @param folderId
	 * @return {@link GetFolder}
	 */
	public GetFolder constructGetFolderById(BaseFolderIdType folderId) {
		NonEmptyArrayOfBaseFolderIdsType foldersArray = new NonEmptyArrayOfBaseFolderIdsType();
		foldersArray.getFolderIdsAndDistinguishedFolderIds().add(folderId);
		return constructGetFolder(foldersArray);
	}

	//================================================================================
    // FindFolder
    //================================================================================	
	/**
	 * Construct a {@link FindFolder} request
	 * @param parent
	 * @param folderShape
	 * @param folderQueryTraversalType
	 * @return {@link FindFolder}
	 */
	public FindFolder constructFindFolder(DistinguishedFolderIdNameType parent, DefaultShapeNamesType folderShape, FolderQueryTraversalType folderQueryTraversalType) {
		return constructFindFolderInternal(parent, folderShape, folderQueryTraversalType, null);
	}
	
	/**
	 * Construct a {@link FindFolder} request
	 * @param parent
	 * @param folderShape
	 * @param folderQueryTraversalType
	 * @param restriction
	 * @return {@link FindFolder}
	 */
	public FindFolder constructFindFolder(DistinguishedFolderIdNameType parent,
			DefaultShapeNamesType folderShape,
			FolderQueryTraversalType folderQueryTraversalType,
			RestrictionType restriction) {
		Validate.notNull(parent, "parent cannot be null");
		Validate.notNull(folderQueryTraversalType, "traversal type cannot be null");
		Validate.notNull(folderShape, "baseShape cannot be null");

		return constructFindFolderInternal(parent, folderShape, folderQueryTraversalType, restriction);
	}
	
	//================================================================================
    // UpdateFolder
    //================================================================================	
	/**
	 * Constructs an {@link UpdateFolder} request for renaming an existing folder
	 * @param newName
	 * @param folderId
	 * @return {@link UpdateFolder}
	 */
	public UpdateFolder constructRenameFolder(String newName, FolderIdType folderId) {
		FolderType folder = new FolderType();
		folder.setDisplayName(newName);
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		path.setFieldURI(UnindexedFieldURIType.FOLDER_DISPLAY_NAME);
		ObjectFactory of = new ObjectFactory();
		return constructUpdateFolderSetField(folder, of.createPath(path), folderId);
	}
	
	//================================================================================
    // EmptyFolder
    //================================================================================
	/**
	 * Construct an {@link EmptyFolder} request
	 * Note this operation cannot be applied to CalendarFolders.
	 * @param deleteSubFolders
	 * @param disposalType
	 * @param folderIds
	 * @return {@link EmptyFolder}
	 */
	public EmptyFolder constructEmptyFolder(boolean deleteSubFolders, DisposalType disposalType, Collection<? extends BaseFolderIdType> folderIds){
		EmptyFolder request = new EmptyFolder();
		request.setDeleteSubFolders(deleteSubFolders);
		request.setDeleteType(disposalType);
		NonEmptyArrayOfBaseFolderIdsType nonEmptyArrayOfBaseFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
		nonEmptyArrayOfBaseFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
		request.setFolderIds(nonEmptyArrayOfBaseFolderIds);
		return request;
	}
	
	//================================================================================
    // DeleteFolder
    //================================================================================
	public DeleteFolder constructDeleteFolder(BaseFolderIdType folderId, DisposalType disposalType) {
		return constructDeleteFolder(Collections.singleton(folderId), disposalType);
	}
	
    //================================================================================
    // CreateItem
    //================================================================================
	/**
	 * Constructs a {@link CreateItem} request specific for creating {@link CalendarItemType} objects
	 * @category CreateItem CalendarItem
	 * 
	 * @param calendarItems
	 * @param sendTo
	 * @param folderId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateCalendarItem(
			Collection<CalendarItemType> calendarItems,
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderId) {
		return constructCreateCalendarItemInternal(calendarItems, sendTo, folderId);
	}
	
	/**
	 * Constructs a {@link CreateItem} request specific for creating
	 * {@link CalendarItemType} objects, within the default calendarFolder and
	 * using {@link ExchangeRequestFactory#getDefaultCalendarCreateSendTo()}
	 * 
	 * @category CreateItem CalendarItem
	 * 
	 * @param calendarItems
	 * @return
	 */
	public CreateItem constructCreateCalendarItem(
			Collection<CalendarItemType> calendarItems) {
		return constructCreateCalendarItemInternal(calendarItems, null, null);
	}

	/**
	 * Constructs a {@link CreateItem} request specific for creating {@link TaskType} objects
	 * @category CreateItem TaskItem
	 * 
	 * @param taskItems
	 * @param folderId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateTaskItem(Collection<TaskType> taskItems, FolderIdType folderId) {
		return constructCreateTaskItem(taskItems, null, folderId);
	}
	
	/**
	 * Constructs a {link CreateItem} request specific for creating {@link MessageType}s (i.e. Email messages)
	 * @param messageItems
	 * @param folderId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateMessageItem(Collection<MessageType> messageItems, FolderIdType folderId){
		return constructCreateMessageItemInternal(messageItems, folderId);
	}
	
	/**
	 * Construct a {@link CreateItem} object which can be used to respond to a calendarInvitation
	 * @param itemId - the itemId of the calendar invite you wish to respond to.
	 * @param response - the type of response you wish to issue.  only {@link ResponseTypeType#ACCEPT}, {@link ResponseTypeType#DECLINE}, and {@link ResponseTypeType#TENTATIVE} are valid.
	 * @return
	 */
	public CreateItem constructCreateItemResponse(ItemIdType itemId, ResponseTypeType response){
		CreateItem request = new CreateItem();
		request.setMessageDisposition(MessageDispositionType.SEND_AND_SAVE_COPY);
		NonEmptyArrayOfAllItemsType arrayOfItems = new NonEmptyArrayOfAllItemsType();
		
		WellKnownResponseObjectType responseType = null;
		switch(response){
			case ACCEPT : responseType = new AcceptItemType();
				break;
			case DECLINE : responseType = new DeclineItemType();
				break;
			case TENTATIVE : responseType = new TentativelyAcceptItemType();
				break;
			default: log.warn(response +" is not a WellKnownResponse type");
				return null;
		}
		responseType.setReferenceItemId(itemId);
		arrayOfItems.getItemsAndMessagesAndCalendarItems().add(responseType);
		request.setItems(arrayOfItems);
		return request;
	}
	
	/**
	 * Construct {@link CreateItem} specific accepting {@link CalendarItemType}s using {@link AcceptItemType}
	 * @category CreateItem AcceptItem
	 * 
	 * @param itemId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateAcceptItemResponse(ItemIdType itemId) {
		return constructCreateItemResponse(itemId, ResponseTypeType.ACCEPT);
	}
	
	/**
	 * Construct {@link CreateItem} specific declining {@link CalendarItemType}s using {@link DeclineItemType}
	 * @category CreateItem AcceptItem
	 * 
	 * @param itemId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateDeclineItemResponse(ItemIdType itemId){
		return constructCreateItemResponse(itemId, ResponseTypeType.DECLINE);
	}
	
	/**
	 * Construct {@link CreateItem} specific tentatively accepting {@link CalendarItemType}s using {@link TentativelyAcceptItemType}
	 * @category CreateItem AcceptItem
	 * 
	 * @param itemId
	 * @return {@link CreateItem}
	 */
	public CreateItem constructCreateTentativeItemResponse(ItemIdType itemId){
		return constructCreateItemResponse(itemId, ResponseTypeType.TENTATIVE);
	}	

	//================================================================================
    // GetItem
    //================================================================================	
	/**
	 * Construct a GetItem request which will only return {@link DefaultShapeNamesType#ID_ONLY} and {@link ExtendedPropertyType}s defined by {@link BaseExchangeRequestFactory#getItemExtendedPropertyPaths()}
	 * @param itemIds
	 * @return
	 */
	public GetItem constructGetItemIds(Collection<ItemIdType> itemIds) {
		Validate.isTrue(!CollectionUtils.isEmpty(itemIds),"itemIds cannot be empty");
		return constructGetItem(itemIds, DefaultShapeNamesType.ID_ONLY);
	}
	/**
	 * Construct a GetItem request which will only return {@link DefaultShapeNamesType#ALL_PROPERTIES} and {@link ExtendedPropertyType}s defined by {@link BaseExchangeRequestFactory#getItemExtendedPropertyPaths()}
	 * @param itemIds
	 * @return
	 */
	public GetItem constructGetItems(Collection<ItemIdType> itemIds) {
		Validate.isTrue(!CollectionUtils.isEmpty(itemIds),"itemIds cannot be empty");
		return constructGetItem(itemIds, DefaultShapeNamesType.ALL_PROPERTIES);
	}

	//================================================================================
    // DeleteItem
    //================================================================================	
	/**
	 * public DeleteItem constructDeleteItem
	 * 
	 * @param itemIds
	 *            - Contains an array of items, occurrence items, and recurring
	 *            master items to delete from a mailbox in the Exchange store.
	 *            The DeleteItem Operation can be performed on any item type
	 * @param disposalType
	 *            - Describes how an item is deleted. This attribute is
	 *            required.
	 * @param sendTo
	 *            - Describes whether a calendar item deletion is communicated
	 *            to attendees. This attribute is required when calendar items
	 *            are deleted. This attribute is optional if non-calendar items
	 *            are deleted.
	 * @param affectedTaskOccurrencesType
	 *            - Describes whether a task instance or a task master is
	 *            deleted by a DeleteItem Operation. This attribute is required
	 *            when tasks are deleted. This attribute is optional when
	 *            non-task items are deleted.
	 * @return
	 */
	public DeleteItem constructDeleteCalendarItems(
			Collection<? extends BaseItemIdType> itemIds,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendTo) {
		Validate.notNull(sendTo, "sendTo must be specified");
		return constructDeleteItem(itemIds, disposalType, sendTo, null);
	}

	/**
	 * Constructs a {@link DeleteItem} request specific for deleting a {@link CalendarItemType}
	 * @param itemId
	 * @param disposalType
	 * @param sendTo
	 * @return
	 */
	public DeleteItem constructDeleteCalendarItem(BaseItemIdType itemId,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendTo) {
		return constructDeleteCalendarItems(Collections.singletonList(itemId),
				disposalType, sendTo);
	}

	/**
	 * Construct a {@link DeleteItem} request specific for deleting a {@link TaskType} 
	 * @param itemIds
	 * @param disposalType
	 * @param affectedTaskOccurrencesType
	 * @return
	 */
	public DeleteItem constructDeleteTaskItems(
			Collection<? extends BaseItemIdType> itemIds,
			DisposalType disposalType,
			AffectedTaskOccurrencesType affectedTaskOccurrencesType) {
		Validate.notNull(affectedTaskOccurrencesType,
				"affectedTaskOccurrencesType must be specified");
		return constructDeleteItem(itemIds, disposalType, null,
				affectedTaskOccurrencesType);
	}
	
	//================================================================================
    // FindItem
    //================================================================================	
	/**
	 * 
	 * FindItem operations
	 * 
	 * 
	 * @param view
	 * @param responseShape
	 * @param traversal
	 *  
	 *            Shallow - Instructs the FindFolder operation to search only
	 *            the identified folder and to return only the folder IDs for
	 *            items that have not been deleted. This is called a shallow
	 *            traversal. Deep - Instructs the FindFolder operation to search
	 *            in all child folders of the identified parent folder and to
	 *            return only the folder IDs for items that have not been
	 *            deleted. This is called a deep traversal. SoftDeleted -
	 *            Instructs the FindFolder operation to perform a shallow
	 *            traversal search for deleted items.
	 * 
	 * @param restriction
	 * @param sortOrderList
	 * @param folderIds
	 * @return
	 */

	/**
	 * Construct a {@link FindItem} request which will return the first {@link ExchangeRequestFactory#getMaxFindItems()} items found.
	 * This method uses an {@link IndexedPageViewType} which supports paging but cannot expand recurrence for you.
	 * 
	 * @param folderIds
	 * @return {@link FindItem}
	 */
	public FindItem constructFindFirstItemIdSet(Collection<FolderIdType> folderIds) {
		return constructFindAllItemIds(INIT_BASE_OFFSET, getMaxFindItems(), folderIds);
	}
	
	/**
	 * Construct a {@link FindItem} request which will return up to {@link ExchangeRequestFactory#getMaxFindItems()} items, beginning with the item at index={@code offset}
	 * If you don't have an offset start with {@link ExchangeRequestFactory#constructFindFirstItemIdSet}
	 * @param offset
	 * @param folderIds
	 * @return {@link FindItem}
	 */
	public FindItem constructFindNextItemIdSet(int offset, Collection<FolderIdType> folderIds) {
		return constructFindAllItemIds(offset, getMaxFindItems(), folderIds);
	}
	
	/**
	 * Construct a {@link FindItem} request which will return a maximum of {@link ExchangeRequestFactory#getMaxFindItems()}, beginning with the item at index={@code offset}
	 * @param offset
	 * @param maxItems
	 * @param folderIds
	 * @return {@link FindItem}
	 */
	public FindItem constructFindAllItemIds(int offset, int maxItems, Collection<FolderIdType> folderIds) {
		//FindAllItems = no restriction
		RestrictionType restriction = null;
		return constructIndexedPageViewFindItemIdsShallow(offset, maxItems, restriction, folderIds);
	}
	
	/**
	 * Constructs a {@link FindItem} request using an {@link IndexedPageViewType}
	 * @param startTime
	 * @param endTime
	 * @param folderIds
	 * @return {@link FindItem}
	 */
	public FindItem constructIndexedPageViewFindItemIdsByDateRange(Date startTime, Date endTime, Collection<FolderIdType> folderIds) {
		Collection<? extends BaseFolderIdType> baseFolderIds = folderIds;
		if(CollectionUtils.isEmpty(baseFolderIds)) {
			DistinguishedFolderIdType distinguishedFolderIdType = new DistinguishedFolderIdType();
			distinguishedFolderIdType.setId(DistinguishedFolderIdNameType.CALENDAR);
			baseFolderIds = Collections.singleton(distinguishedFolderIdType);
		}
		
		RestrictionType restriction = constructFindCalendarItemsByDateRangeRestriction(startTime, endTime);
		return constructIndexedPageViewFindFirstItemIdsShallow(restriction, baseFolderIds);
	}
	
	/**
	 * Constructs a {@link FindItem} request using a {@link CalendarViewType}
	 * @param startTime
	 * @param endTime
	 * @param folderIds
	 * @return {@link FindItem}
	 */
	public FindItem constructCalendarViewFindCalendarItemIdsByDateRange(Date startTime, Date endTime, Collection<FolderIdType> folderIds) {
		Collection<? extends BaseFolderIdType> baseFolderIds = folderIds;
		if(CollectionUtils.isEmpty(baseFolderIds)) {
			DistinguishedFolderIdType distinguishedFolderIdType = new DistinguishedFolderIdType();
			distinguishedFolderIdType.setId(DistinguishedFolderIdNameType.CALENDAR);
			baseFolderIds = Collections.singleton(distinguishedFolderIdType);
		}
		return constructCalendarViewFindItemIdsByDateRange(startTime, endTime, baseFolderIds);
	}
}
