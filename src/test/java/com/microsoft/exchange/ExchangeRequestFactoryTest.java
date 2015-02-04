/**
 * 
 */
package com.microsoft.exchange;



import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Set;

import javax.xml.bind.JAXBElement;

import org.apache.commons.lang.time.DateUtils;
import org.junit.Test;
import org.springframework.dao.support.DataAccessUtils;

import com.microsoft.exchange.impl.ExchangeRequestFactoryUtils;
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
import com.microsoft.exchange.messages.UpdateItem;
import com.microsoft.exchange.types.AcceptItemType;
import com.microsoft.exchange.types.AndType;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.BaseFolderType;
import com.microsoft.exchange.types.BasePathToElementType;
import com.microsoft.exchange.types.BodyTypeResponseType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.ConstantValueType;
import com.microsoft.exchange.types.DeclineItemType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.DistinguishedFolderIdType;
import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.FolderChangeDescriptionType;
import com.microsoft.exchange.types.FolderChangeType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.FolderResponseShapeType;
import com.microsoft.exchange.types.FreeBusyViewOptions;
import com.microsoft.exchange.types.IndexBasePointType;
import com.microsoft.exchange.types.IndexedPageViewType;
import com.microsoft.exchange.types.IsGreaterThanOrEqualToType;
import com.microsoft.exchange.types.IsLessThanType;
import com.microsoft.exchange.types.ItemChangeType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.ItemQueryTraversalType;
import com.microsoft.exchange.types.ItemResponseShapeType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.Mailbox;
import com.microsoft.exchange.types.MailboxData;
import com.microsoft.exchange.types.MessageDispositionType;
import com.microsoft.exchange.types.MessageType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseFolderIdsType;
import com.microsoft.exchange.types.NonEmptyArrayOfFolderChangesType;
import com.microsoft.exchange.types.NonEmptyArrayOfFoldersType;
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangesType;
import com.microsoft.exchange.types.NonEmptyArrayOfTimeZoneIdType;
import com.microsoft.exchange.types.PathToUnindexedFieldType;
import com.microsoft.exchange.types.ResponseTypeType;
import com.microsoft.exchange.types.RestrictionType;
import com.microsoft.exchange.types.SearchExpressionType;
import com.microsoft.exchange.types.SearchFolderType;
import com.microsoft.exchange.types.SetFolderFieldType;
import com.microsoft.exchange.types.SetItemFieldType;
import com.microsoft.exchange.types.SuggestionsViewOptions;
import com.microsoft.exchange.types.TargetFolderIdType;
import com.microsoft.exchange.types.TaskType;
import com.microsoft.exchange.types.TasksFolderType;
import com.microsoft.exchange.types.TentativelyAcceptItemType;
import com.microsoft.exchange.types.TimeZone;
import com.microsoft.exchange.types.UnindexedFieldURIType;
import com.microsoft.exchange.types.WellKnownResponseObjectType;


/**
 * Tests for {@link ExchangeRequestFactory}.
 * 
 * @author Collin Cudd
 */
public class ExchangeRequestFactoryTest extends ExchangeRequestFactoryUtils{
	ExchangeRequestFactory factory = new ExchangeRequestFactory();
	/**
	 * @return the factory
	 */
	@Override
	public ExchangeRequestFactory getFactory() {
		return factory;
	}
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCalendarViewFindCalendarItemIdsByDateRange(Date, Date, java.util.Collection)}
	 */
	@Test
	public void constructCalendarViewFindCalendarItemIdsByDateRange_control(){
		Date startTime = new Date();
		Date endTime = DateUtils.addDays(startTime, 7);
		FindItem request = factory.constructCalendarViewFindCalendarItemIdsByDateRange(startTime, endTime, null);
		assertNotNull(request);
		
		//verify the calendar view
		CalendarViewType view = findItemIsCalendarView(request);
		assertEquals(ExchangeDateUtils.convertDateToXMLGregorianCalendar(startTime), view.getStartDate());
		assertEquals(ExchangeDateUtils.convertDateToXMLGregorianCalendar(endTime),view.getEndDate());
		assertEquals(factory.getMaxFindItems(), view.getMaxEntriesReturned().intValue());
		
		assertNull(request.getDistinguishedGroupBy());
		assertNull(request.getGroupBy());
		
		//verify the desitnation is the primary calendar folder
		NonEmptyArrayOfBaseFolderIdsType parentFolderIds = request.getParentFolderIds();
		assertNotNull(parentFolderIds);
		List<BaseFolderIdType> folderIds = parentFolderIds.getFolderIdsAndDistinguishedFolderIds();
		assertNotNull(folderIds);
		assertEquals(1, folderIds.size());
		assertEquals(ExchangeRequestFactory.getPrimaryCalendarDistinguishedFolderId(), folderIds.get(0));

		//verify the response shape
		assertTrue(checkResponseShape(request.getItemShape(), DefaultShapeNamesType.ID_ONLY, BodyTypeResponseType.TEXT));
		
		//restrictions, queries and sort orders CANNOT be applied to calendarViews
		assertNull(request.getQueryString());
		assertNull(request.getRestriction());
		assertNull(request.getSortOrder());
		
		//verify the traversal type
		assertTrue(checkTraversal(request.getTraversal(), ItemQueryTraversalType.SHALLOW));
	}
	
	//================================================================================
    // Tests for methods which produce CreateItem for WellKnownResponseType
    //================================================================================	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateItemResponse(ItemIdType, ResponseTypeType)}
	 */
	@Test
	public void constructCreateItemResponse_control(){
		ItemIdType itemId = new ItemIdType();
		assertEquals(factory.constructCreateAcceptItemResponse(itemId), factory.constructCreateItemResponse(itemId, ResponseTypeType.ACCEPT));
		assertEquals(factory.constructCreateDeclineItemResponse(itemId), factory.constructCreateItemResponse(itemId, ResponseTypeType.DECLINE));
		assertEquals(factory.constructCreateTentativeItemResponse(itemId), factory.constructCreateItemResponse(itemId, ResponseTypeType.TENTATIVE));
		assertNull(factory.constructCreateItemResponse(itemId, ResponseTypeType.NO_RESPONSE_RECEIVED));
		assertNull(factory.constructCreateItemResponse(itemId, ResponseTypeType.ORGANIZER));
		assertNull(factory.constructCreateItemResponse(itemId, ResponseTypeType.UNKNOWN));
	}
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateAcceptItemResponse(ItemIdType)}.
	 */
	@Test
	public void constructCreateAcceptItemResponse_control(){
		ItemIdType itemId = new ItemIdType();
		CreateItem request = factory.constructCreateAcceptItemResponse(itemId);
		WellKnownResponseObjectType item = createItemIsWellKnownResponseType(request);
		assertNotNull(item);
		assertTrue(item instanceof AcceptItemType);
		AcceptItemType accept = (AcceptItemType) item;
		assertEquals(itemId, accept.getReferenceItemId());
		assertEquals(MessageDispositionType.SEND_AND_SAVE_COPY,request.getMessageDisposition());
	}
	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateDeclineItemResponse}
	 */
	@Test
	public void constructCreateDeclineItemResponse_control(){
		ItemIdType itemId = new ItemIdType();
		CreateItem request = factory.constructCreateDeclineItemResponse(itemId);
		WellKnownResponseObjectType item = createItemIsWellKnownResponseType(request);
		assertNotNull(item);
		assertTrue(item instanceof DeclineItemType);
		DeclineItemType decline = (DeclineItemType) item;
		assertEquals(itemId, decline.getReferenceItemId());
		assertEquals(MessageDispositionType.SEND_AND_SAVE_COPY,request.getMessageDisposition());
	}
	
	/**
	 * Control test {@link ExchangeRequestFactory#constructCreateTentativeItemResponse(ItemIdType)}
	 */
	@Test
	public void constructCreateTentativeItemResponse_control(){
		ItemIdType itemId = new ItemIdType();
		CreateItem request = factory.constructCreateTentativeItemResponse(itemId);
		WellKnownResponseObjectType item = createItemIsWellKnownResponseType(request);
		assertNotNull(item);
		assertTrue(item instanceof TentativelyAcceptItemType);
		TentativelyAcceptItemType tentative = (TentativelyAcceptItemType) item;
		assertEquals(itemId, tentative.getReferenceItemId());
		assertEquals(MessageDispositionType.SEND_AND_SAVE_COPY,request.getMessageDisposition());
	}

	//================================================================================
    // Tests for methods which produce CreateItem for FolderTypes
    //================================================================================	
	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateSearchFolder(String, com.microsoft.exchange.types.DistinguishedFolderIdNameType, com.microsoft.exchange.types.RestrictionType)}
	 */
	@Test
	public void constructCreateSearchFolder_control(){
		String displayName = "displayname";
		DistinguishedFolderIdNameType searchRoot = DistinguishedFolderIdNameType.CALENDAR;
		RestrictionType restriction = new RestrictionType();
		CreateFolder request = factory.constructCreateSearchFolder(displayName, searchRoot, restriction);
		
		assertNotNull(request);
		NonEmptyArrayOfFoldersType folders = request.getFolders();
		assertNotNull(folders);
		List<BaseFolderType> folderList = folders.getFoldersAndCalendarFoldersAndContactsFolders();
		assertEquals(1, folderList.size());
		BaseFolderType baseFolderType = folderList.get(0);
		assertTrue(baseFolderType instanceof SearchFolderType);
		assertEquals(displayName, baseFolderType.getDisplayName());
		
		TargetFolderIdType parentFolderId = request.getParentFolderId();
		assertNotNull(parentFolderId);
		assertEquals(ExchangeRequestFactory.getParentTargetFolderId(DistinguishedFolderIdNameType.SEARCHFOLDERS), parentFolderId);
		
		//TODO this is incomplete.
	}
	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateCalendarFolder(String, Collection)}.
	 */
	@Test
	public void constructCreateCalendarFolder_control(){
		String displayName = "displayname";
		CreateFolder request = factory.constructCreateCalendarFolder(displayName, null);
		assertNotNull(request);
		TargetFolderIdType parentFolderId = request.getParentFolderId();
		assertNotNull(parentFolderId);
		assertEquals(ExchangeRequestFactory.getPrimaryCalendarDistinguishedFolderId(), parentFolderId.getDistinguishedFolderId());
	
		NonEmptyArrayOfFoldersType folders = request.getFolders();
		assertNotNull(folders);
		List<BaseFolderType> folderList = folders.getFoldersAndCalendarFoldersAndContactsFolders();
		assertEquals(1, folderList.size());
		BaseFolderType baseFolderType = folderList.get(0);
		assertTrue(baseFolderType instanceof CalendarFolderType);
		assertEquals(displayName, baseFolderType.getDisplayName());
	}
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateTaskFolder(String, Collection)}
	 */
	@Test
	public void constructCreateTaskFolder_control(){
		String displayName = "displayname";
		Collection<ExtendedPropertyType> exProps = new ArrayList<ExtendedPropertyType>();
		CreateFolder request = factory.constructCreateTaskFolder(displayName, exProps);
		assertNotNull(request);
		TargetFolderIdType parentFolderId = request.getParentFolderId();
		assertNotNull(parentFolderId);
		assertEquals(ExchangeRequestFactory.getPrimaryTasksDistinguishedFolderId(), parentFolderId.getDistinguishedFolderId());
	
		NonEmptyArrayOfFoldersType folders = request.getFolders();
		assertNotNull(folders);
		List<BaseFolderType> folderList = folders.getFoldersAndCalendarFoldersAndContactsFolders();
		assertEquals(1, folderList.size());
		BaseFolderType baseFolderType = folderList.get(0);
		assertTrue(baseFolderType instanceof TasksFolderType);
		assertEquals(displayName, baseFolderType.getDisplayName());		
	}
	
	//================================================================================
    // Tests for methods which produce CreateItem for ItemTypes
    //================================================================================		
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateCalendarItem(Collection)}
	 */
	@Test
	public void constructCreateCalendarItem_control(){
		CalendarItemType calendarItem = new CalendarItemType();
		Collection<CalendarItemType> calendarItems = Collections.<CalendarItemType>singleton(calendarItem);
		CreateItem request = factory.constructCreateCalendarItem(calendarItems);
		assertNotNull(request);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		
		assertTrue(item instanceof CalendarItemType);
		assertEquals(calendarItem, item);
		
		assertNull(request.getMessageDisposition());
		assertEquals(factory.getSendMeetingInvitations(), request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertNull(savedItemFolderId.getFolderId());
		assertEquals(ExchangeRequestFactory.getPrimaryCalendarDistinguishedFolderId(), savedItemFolderId.getDistinguishedFolderId());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructCreateCalendarItem(Collection, FolderIdType))}
	 */
	@Test
	public void constructCreateCalendarItem_with_folderId(){
		CalendarItemType calendarItem = new CalendarItemType();
		FolderIdType folderId = createFolderId("calendarFolderId",null);
		
		Collection<CalendarItemType> calendarItems = Collections.<CalendarItemType>singleton(calendarItem);
		CreateItem request = factory.constructCreateCalendarItem(calendarItems, folderId);
		assertNotNull(request);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		
		assertTrue(item instanceof CalendarItemType);
		assertEquals(calendarItem, item);
		
		assertNull(request.getMessageDisposition());
		assertEquals(factory.getSendMeetingInvitations(), request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertNull(savedItemFolderId.getDistinguishedFolderId());
		assertEquals(folderId.getId(), savedItemFolderId.getFolderId().getId());
	}
	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateMessageItem(Collection, com.microsoft.exchange.types.FolderIdType)}
	 */
	@Test
	public void constructCreateMessageItem_control(){
		MessageType message = new MessageType();
		Collection<MessageType> messageItems = Collections.singleton(message);
		FolderIdType folderId = null;
		CreateItem request = factory.constructCreateMessageItem(messageItems, folderId);
		assertNotNull(request);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		
		assertTrue(item instanceof MessageType);
		assertEquals(message, item);
		
		MessageDispositionType messageDisposition = request.getMessageDisposition();
		assertNotNull(messageDisposition);
		assertEquals(factory.getMessageDisposition(), messageDisposition);
		
		assertNull(request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertEquals(ExchangeRequestFactory.getPrimaryMessageDistinguishedFolderId(), savedItemFolderId.getDistinguishedFolderId());
	}
	
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateMessageItem(Collection, com.microsoft.exchange.types.FolderIdType)}
	 */
	@Test
	public void constructCreateMessageItem_with_foldlerId(){
		MessageType message = new MessageType();
		Collection<MessageType> messageItems = Collections.singleton(message);
		FolderIdType folderId = createFolderId("foo",null);
		CreateItem request = factory.constructCreateMessageItem(messageItems, folderId);
		assertNotNull(request);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		
		assertTrue(item instanceof MessageType);
		assertEquals(message, item);
		
		MessageDispositionType messageDisposition = request.getMessageDisposition();
		assertNotNull(messageDisposition);
		assertEquals(factory.getMessageDisposition(), messageDisposition);
		
		assertNull(request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertNull(savedItemFolderId.getDistinguishedFolderId());
		assertEquals(folderId.getId(), savedItemFolderId.getFolderId().getId());
	}

	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateTaskItem(Collection)}
	 */
	@Test
	public void constructCreateTaskItem_control(){
		TaskType task = new TaskType();
		Collection<TaskType> taskItems = Collections.singleton(task);
		
		FolderIdType folderId = new FolderIdType();
		CreateItem request = factory.constructCreateTaskItem(taskItems, folderId);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		assertTrue(item instanceof TaskType);
		assertEquals(task, item);
	
		assertNull(request.getMessageDisposition());
		assertNull(request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertNotNull(savedItemFolderId.getDistinguishedFolderId());
		assertEquals(ExchangeRequestFactory.getPrimaryTasksDistinguishedFolderId(), savedItemFolderId.getDistinguishedFolderId());
		assertNull(savedItemFolderId.getFolderId());
	}
	/**
	 * Control test for {@link ExchangeRequestFactory#constructCreateTaskItem(Collection, FolderIdType)}
	 */
	@Test
	public void constructCreateTaskItem_with_folderId(){
		TaskType task = new TaskType();
		Collection<TaskType> taskItems = Collections.singleton(task);
		
		FolderIdType folderId = new FolderIdType();
		folderId.setId("taskFolderId");
		CreateItem request = factory.constructCreateTaskItem(taskItems, folderId);
		
		ItemType item = createItemContainsSingleItem(request);
		assertNotNull(item);
		assertTrue(item instanceof TaskType);
		assertEquals(task, item);
	
		assertNull(request.getMessageDisposition());
		assertNull(request.getSendMeetingInvitations());
		
		TargetFolderIdType savedItemFolderId = request.getSavedItemFolderId();
		assertNotNull(savedItemFolderId);
		assertNull(savedItemFolderId.getDistinguishedFolderId());
		assertEquals(folderId, savedItemFolderId.getFolderId());
	}
	
	//================================================================================
    // Tests for methods used for Deletion
    //================================================================================		

	/**
	 * Test for {@link ExchangeRequestFactory#constructDeleteCalendarItems(Collection)}
	 */
	@Test
	public void constructDeleteCalendarItems_control(){
		ItemIdType itemId = new ItemIdType();
		itemId.setId("itemid");
		itemId.setChangeKey("changeKey");

		DeleteItem request = factory.constructDeleteCalendarItems(Collections.singleton(itemId));
		assertNotNull(request);

		assertNull(request.getAffectedTaskOccurrences());
		assertNotNull(request.getDeleteType());
		assertEquals(factory.getItemDisposalType(),request.getDeleteType());
		
		assertNotNull(request.getItemIds());
		assertEquals(itemId,containsSingleItemId(request.getItemIds()));
		
		assertNotNull(request.getSendMeetingCancellations());
		assertEquals(factory.getSendMeetingInvitations(), request.getSendMeetingCancellations());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructDeleteFolder(BaseFolderIdType)}
	 */
	@Test
	public void constructDeleteFolder_control(){
		FolderIdType folderId = createFolderId();
		
		DeleteFolder request = factory.constructDeleteFolder(folderId);
		assertNotNull(request);
		
		assertNotNull(request.getDeleteType());
		assertEquals(factory.getFolderDisposalType(), request.getDeleteType());
		
		assertNotNull(request.getFolderIds());
		assertEquals(folderId, containsSingleFolderId(request.getFolderIds()));
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructDeleteCalendarItems(Collection)}
	 */
	@Test
	public void constructDeleteTaskItems_control(){
		ItemIdType itemId = new ItemIdType();
		itemId.setId("itemid");
		itemId.setChangeKey("changeKey");
		
		DeleteItem request = factory.constructDeleteTaskItems(Collections.singleton(itemId));
		assertNotNull(request);
		
		assertNotNull(request.getAffectedTaskOccurrences());
		assertEquals(factory.getAffectedTaskOccurrencesType(), request.getAffectedTaskOccurrences());
		
		assertNotNull(request.getDeleteType());
		assertEquals(factory.getItemDisposalType(), request.getDeleteType());
		
		assertNotNull(request.getItemIds());
		assertEquals(itemId, containsSingleItemId(request.getItemIds()));
		
		assertNull(request.getSendMeetingCancellations());
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructEmptyFolder(boolean, Collection)}
	 */
	@Test
	public void constructEmptyFolder_test(){
		FolderIdType folderId = createFolderId();
		
		Collection<? extends BaseFolderIdType> folderIds = Collections.singleton(folderId);
		boolean deleteSubFolders = true;
		EmptyFolder request = factory.constructEmptyFolder(deleteSubFolders, folderIds);
		assertNotNull(request);
		
		assertNotNull(request.getDeleteType());
		assertEquals(factory.getFolderDisposalType(), request.getDeleteType());
		
		assertNotNull(request.getFolderIds());
		assertEquals(folderId, containsSingleFolderId(request.getFolderIds()));
		
		assertEquals(deleteSubFolders, request.isDeleteSubFolders());		
	}
	
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructFindFirstItemIdSet(Collection)}
	 */
	@Test
	public void constructFindFirstItemIdSet_test(){
		FolderIdType folderId = createFolderId();
		
		Set<FolderIdType> folderIds = Collections.singleton(folderId);
		FindItem request = factory.constructFindFirstItemIdSet(folderIds);
		assertNotNull(request);
		
		IndexedPageViewType view = findItemIsIndexedPageView(request);
		assertNotNull(view);
		
		//the view should start at the begin
		assertEquals(0, view.getOffset());
		assertEquals(IndexBasePointType.BEGINNING, view.getBasePoint());
		assertEquals(factory.getMaxFindItems(),view.getMaxEntriesReturned().intValue());
		
		assertNotNull(request.getParentFolderIds());
		assertEquals(folderId,containsSingleFolderId(request.getParentFolderIds()));
		
		assertNotNull(request.getItemShape());
		assertEquals(DefaultShapeNamesType.ID_ONLY, request.getItemShape().getBaseShape());
		
		//no search restrictions we are getting all limited only by maxFind
		assertNull(request.getRestriction());
		
		assertNotNull(request.getTraversal());
		//we are only looking for items in the folderIds provided (dont search recursively)
		assertEquals(ItemQueryTraversalType.SHALLOW, request.getTraversal());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructFindNextItemIdSet(int, Collection)}
	 */
	@Test
	public void constructFindNextItemIdSet(){
		int offset = 1210;
		FolderIdType folderId = createFolderId();
		
		FindItem request = factory.constructFindNextItemIdSet(offset, Collections.singleton(folderId));
		assertNotNull(request);
		
		IndexedPageViewType view = findItemIsIndexedPageView(request);
		assertNotNull(view);
		
		//the view should start at the begin
		assertEquals(offset, view.getOffset());
		assertEquals(IndexBasePointType.BEGINNING, view.getBasePoint());
		assertEquals(factory.getMaxFindItems(),view.getMaxEntriesReturned().intValue());
		
		assertNotNull(request.getParentFolderIds());
		assertEquals(folderId,containsSingleFolderId(request.getParentFolderIds()));
		
		ItemResponseShapeType itemShape = request.getItemShape();
		assertNotNull(itemShape);
		//looking for itemIds only
		assertEquals(DefaultShapeNamesType.ID_ONLY, itemShape.getBaseShape());
		assertNull(itemShape.getAdditionalProperties());
		
		//no search restrictions we are getting all limited only by maxFind
		assertNull(request.getRestriction());
		
		assertNotNull(request.getTraversal());
		//we are only looking for items in the folderIds provided (dont search recursively)
		assertEquals(ItemQueryTraversalType.SHALLOW, request.getTraversal());
		
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructFindFolder(DistinguishedFolderIdNameType)}
	 */
	@Test
	public void constructFindFolder(){
		DistinguishedFolderIdNameType parent = DistinguishedFolderIdNameType.CALENDAR;
		FindFolder request = factory.constructFindFolder(parent);
		assertNotNull(request);
		
		//we want all props
		assertNotNull(request.getFolderShape());
		assertEquals(DefaultShapeNamesType.ALL_PROPERTIES,request.getFolderShape().getBaseShape());
		
		//we are expecting an indexed view
		IndexedPageViewType indexedPageFolderView = request.getIndexedPageFolderView();
		assertNotNull(indexedPageFolderView);
		//not a fractional view
		assertNull(request.getFractionalPageFolderView());
		
		//the view should start at the begin
		assertEquals(0, indexedPageFolderView.getOffset());
		assertEquals(IndexBasePointType.BEGINNING, indexedPageFolderView.getBasePoint());
		assertEquals(factory.getMaxFindItems(),indexedPageFolderView.getMaxEntriesReturned().intValue());

		
		assertNotNull(request.getParentFolderIds());
		assertEquals(ExchangeRequestFactory.getPrimaryCalendarDistinguishedFolderId(), containsSingleFolderId(request.getParentFolderIds()));
		
		assertNull(request.getRestriction());

		assertNotNull(request.getTraversal());
		assertEquals(factory.getFolderQueryTraversal(), request.getTraversal());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetFolderByDistinguishedName(DistinguishedFolderIdNameType)}
	 */
	@Test
	public void constructGetFolderByDistinguishedName(){
		DistinguishedFolderIdNameType name = DistinguishedFolderIdNameType.INBOX;
		GetFolder request = factory.constructGetFolderByDistinguishedName(name);
		assertNotNull(request);
		
		FolderResponseShapeType folderShape = request.getFolderShape();
		assertNotNull(folderShape);
		assertNull(folderShape.getAdditionalProperties());
		assertNotNull(folderShape.getBaseShape());
		assertEquals(DefaultShapeNamesType.ALL_PROPERTIES, folderShape.getBaseShape());
		
		assertNotNull(request.getFolderIds());
		assertEquals(ExchangeRequestFactory.getPrimaryMessageDistinguishedFolderId(), containsSingleFolderId(request.getFolderIds()));
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetFolderById(BaseFolderIdType)}
	 */
	@Test
	public void constructGetFolderById(){
		BaseFolderIdType folderId = createFolderId("123", "abc");
		GetFolder request = factory.constructGetFolderById(folderId);
		assertNotNull(request);
		
		FolderResponseShapeType folderShape = request.getFolderShape();
		assertNotNull(folderShape);
		assertNull(folderShape.getAdditionalProperties());
		assertNotNull(folderShape.getBaseShape());
		assertEquals(DefaultShapeNamesType.ALL_PROPERTIES, folderShape.getBaseShape());
		
		assertNotNull(request.getFolderIds());
		assertEquals(folderId, containsSingleFolderId(request.getFolderIds()));
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetItemIds(Collection)}
	 */
	@Test
	public void constructGetItemIds(){
		ItemIdType createItemId = createItemId();
		Collection<ItemIdType> itemIds = Collections.singleton(createItemId);
		GetItem request = factory.constructGetItemIds(itemIds);
		assertNotNull(request);
		
		assertNotNull(request.getItemIds());
		assertEquals(createItemId,containsSingleItemId(request.getItemIds()));
		
		ItemResponseShapeType shape = request.getItemShape();
		assertNotNull(shape);
		assertNull(shape.getAdditionalProperties());
		
		assertNotNull(shape.getBaseShape());
		assertEquals(DefaultShapeNamesType.ID_ONLY, shape.getBaseShape());
		
		BodyTypeResponseType body = shape.getBodyType();
		assertNotNull(body);
		assertEquals(BodyTypeResponseType.TEXT, body);
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetItems(Collection)} 
	 */
	@Test
	public void constructGetItems(){
		ItemIdType createItemId = createItemId();
		Collection<ItemIdType> itemIds = Collections.singleton(createItemId);
		GetItem request = factory.constructGetItems(itemIds);
		assertNotNull(request);
		
		assertNotNull(request.getItemIds());
		assertEquals(createItemId, containsSingleItemId(request.getItemIds()));
		
		ItemResponseShapeType shape = request.getItemShape();
		assertNotNull(shape);
		assertNull(shape.getAdditionalProperties());
		
		assertNotNull(shape.getBaseShape());
		assertEquals(DefaultShapeNamesType.ALL_PROPERTIES, shape.getBaseShape());
		
		BodyTypeResponseType body = shape.getBodyType();
		assertNotNull(body);
		assertEquals(BodyTypeResponseType.TEXT, body);
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetServerTimeZones(String, boolean)}
	 */
	@Test
	public void constructGetServerTimeZones(){
		boolean returnFullTimeZoneData = false;
		String tzid = "tzid";
		GetServerTimeZones request = factory.constructGetServerTimeZones(tzid, returnFullTimeZoneData);
		assertNotNull(request);
		
		NonEmptyArrayOfTimeZoneIdType ids = request.getIds();
		assertNotNull(ids);
		List<String> tzids = ids.getIds();
		assertEquals(tzid, DataAccessUtils.singleResult(tzids));
		
		assertEquals(returnFullTimeZoneData, request.isReturnFullTimeZoneData());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetUserAvailabilityRequest(Collection, FreeBusyViewOptions, SuggestionsViewOptions, TimeZone)} 
	 */
	@Test
	public void constructGetUserAvailabilityRequest(){
		MailboxData mailbox = new MailboxData();
		Mailbox value =  new Mailbox();
		value.setName("name");
		value.setAddress("addr");
		mailbox.setEmail(value );
		Collection<? extends MailboxData> mailboxData = Collections.singleton(mailbox);
		FreeBusyViewOptions originalFreeBusyOption = new FreeBusyViewOptions();
		SuggestionsViewOptions suggestionsView = null;
		TimeZone timeZone = null;
		GetUserAvailabilityRequest request = factory.constructGetUserAvailabilityRequest(mailboxData, originalFreeBusyOption, suggestionsView, timeZone);
		assertNotNull(request);
		
		assertNotNull(request.getFreeBusyViewOptions());
		assertEquals(originalFreeBusyOption, request.getFreeBusyViewOptions());
		
		assertNotNull(request.getMailboxDataArray());
		//assertEquals(mailbox, Collections.singleton(request.getMailboxDataArray().getMailboxDatas()));
		
		//TODO
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructGetUserConfiguration(String, DistinguishedFolderIdType)}.
	 */
	@Test
	public void constructGetUserConfiguration(){
		String name = null;
		DistinguishedFolderIdType distinguishedFolderIdType = null;
		GetUserConfiguration request = factory.constructGetUserConfiguration(name, distinguishedFolderIdType);
		assertNotNull(request);
		//TODO
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructIndexedPageViewFindItemIdsByDateRange(Date, Date, Collection)} 
	 */
	@Test
	public void constructIndexedPageViewFindItemIdsByDateRange(){
		Date startTime = new Date();
		Date endTime = DateUtils.addDays(startTime, 7);
		Collection<FolderIdType> folderIds = Collections.singleton(createFolderId());
		FindItem request = factory.constructIndexedPageViewFindItemIdsByDateRange(startTime, endTime, folderIds);
		assertNotNull(request);
		//we are expecting a IndexedPageViewType
		IndexedPageViewType view = findItemIsIndexedPageView(request);
		assertNotNull(view);
		
		//the IndexedPageViewType should start at the beginning, i.e. 0, and return at most factory.getMaxFindItems()
		assertTrue(indexedPageViewTypeBasePointIsBeginning(view));	
		
		//the itemShape must be present, must be ID only for baseShape and Text for bodyResponseType.
		assertTrue(checkResponseShape(request.getItemShape(), DefaultShapeNamesType.ID_ONLY, BodyTypeResponseType.TEXT));
		
		//we want a SHALLOW traversal type, we don't need to search deleted or associated to find calendarItemIds
		assertTrue(checkTraversal(request.getTraversal(), ItemQueryTraversalType.SHALLOW));
		
		//we want Restrictions
		RestrictionType restriction = request.getRestriction();
		assertNotNull(restriction);
		assertNotNull(restriction.getSearchExpression());
		SearchExpressionType searchExpression = restriction.getSearchExpression().getValue();
		assertNotNull(searchExpression);
		assertTrue(searchExpression instanceof AndType);
		AndType andType = (AndType) searchExpression;
		List<JAXBElement<? extends SearchExpressionType>> searchExpressions = andType.getSearchExpressions();
		assertEquals(2, searchExpressions.size());
		for(JAXBElement<? extends SearchExpressionType> element : searchExpressions){
			SearchExpressionType se = element.getValue();
			if(se instanceof IsGreaterThanOrEqualToType){
				IsGreaterThanOrEqualToType greaterEqual = (IsGreaterThanOrEqualToType) se;
				ConstantValueType constant = greaterEqual.getFieldURIOrConstant().getConstant();
				assertEquals(ExchangeDateUtils.convertDateToXMLGregorianCalendar(startTime).toXMLFormat(), constant.getValue());
			}else if(se instanceof IsLessThanType){
				IsLessThanType less = (IsLessThanType) se;
				ConstantValueType constant = less.getFieldURIOrConstant().getConstant();
				assertEquals(ExchangeDateUtils.convertDateToXMLGregorianCalendar(endTime).toXMLFormat(), constant.getValue());
			}else{
				fail("unexpected search expression type: "+se.getClass());
			}
		}
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructRenameFolder(String, FolderIdType)}
	 */
	@Test
	public void constructRenameFolder(){
		String newName = "newname";
		FolderIdType folderId = createFolderId();
		UpdateFolder request = factory.constructRenameFolder(newName, folderId);
		assertNotNull(request);
		
		NonEmptyArrayOfFolderChangesType folderChanges = request.getFolderChanges();
		assertNotNull(folderChanges);
		
		FolderChangeType change = DataAccessUtils.singleResult(folderChanges.getFolderChanges());
		assertNotNull(change);
		assertNotNull(change.getFolderId());
		assertEquals(folderId,change.getFolderId());
		
		assertNotNull(change.getUpdates());
		FolderChangeDescriptionType changeDesc = DataAccessUtils.singleResult(change.getUpdates().getAppendToFolderFieldsAndSetFolderFieldsAndDeleteFolderFields());
		assertNotNull(changeDesc);
		
		BasePathToElementType basePath = changeDesc.getPath().getValue();
		assertNotNull(basePath);
		assertTrue(basePath instanceof PathToUnindexedFieldType);
		PathToUnindexedFieldType path = (PathToUnindexedFieldType) basePath;
		
		assertTrue(changeDesc instanceof SetFolderFieldType);
		SetFolderFieldType setFolderField = (SetFolderFieldType) changeDesc;
		assertNotNull(setFolderField.getFolder());
		assertEquals(newName, setFolderField.getFolder().getDisplayName());
		assertEquals(UnindexedFieldURIType.FOLDER_DISPLAY_NAME, path.getFieldURI());
		
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructResolveNames(String)} 
	 */
	@Test
	public void constructResolveNames(){
		String alias = "alias";
		ResolveNames request = factory.constructResolveNames(alias);
		assertNotNull(request);
		
		DefaultShapeNamesType contactDataShape = request.getContactDataShape();
		assertNotNull(contactDataShape);
		assertEquals(DefaultShapeNamesType.ALL_PROPERTIES, contactDataShape);
		
		assertNotNull(request.getSearchScope());
		assertEquals(factory.getResolveNamesSearchScope(), request.getSearchScope());
		
		assertTrue(request.isReturnFullContactData());
		
		assertNotNull(request.getUnresolvedEntry());
		assertEquals(alias, request.getUnresolvedEntry());
		
		assertNull(request.getParentFolderIds());
	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructSetCalendarItemExtendedProperty(CalendarItemType, ExtendedPropertyType)}
	 */
	@Test
	public void constructSetCalendarItemExtendedProperty(){
		CalendarItemType item = new CalendarItemType();
		ExtendedPropertyType exprop = new ExtendedPropertyType();
		SetItemFieldType request = factory.constructSetCalendarItemExtendedProperty(item, exprop);
		assertNotNull(request);
		//TODO
	}
	
	/**
	 * Test for Test for {@link ExchangeRequestFactory#constructSetCalendarItemRequiredAttendees(CalendarItemType)} 
	 */
	@Test
	public void constructSetCalendarItemRequiredAttendees(){
		
		CalendarItemType item = new CalendarItemType();
		SetItemFieldType request = factory.constructSetCalendarItemRequiredAttendees(item);
		
		//TODO
	}
	/**
	 * Test for {@link ExchangeRequestFactory#constructSetCalendarItemSubject(CalendarItemType)}
	 */
	@Test
	public void constructSetCalendarItemSubject(){
		CalendarItemType original = new CalendarItemType();
		original.setStart(ExchangeDateUtils.convertDateToXMLGregorianCalendar(new Date()));
		original.setSubject("subject");
		
		SetItemFieldType setSubject = factory.constructSetCalendarItemSubject(original);
		assertNotNull(setSubject);
		
		//the calendarItem within the change field has only the property we intend to change, no other properties are set.
		CalendarItemType changeItem = setSubject.getCalendarItem();
		assertNotNull(changeItem);
		assertNotNull(changeItem.getSubject());
		assertEquals(original.getSubject(), changeItem.getSubject());
		assertNull(changeItem.getStart());
		
		JAXBElement<? extends BasePathToElementType> path = setSubject.getPath();
		assertNotNull(path);
		BasePathToElementType basePath = path.getValue();
		assertNotNull(basePath);
		assertTrue(basePath instanceof PathToUnindexedFieldType);
		PathToUnindexedFieldType unindexedPath = (PathToUnindexedFieldType) basePath;
		assertNotNull(unindexedPath);
		assertEquals(UnindexedFieldURIType.ITEM_SUBJECT, unindexedPath.getFieldURI());

	}
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructSetCalendarItemLegacyFreeBusy(CalendarItemType)}
	 */
	@Test
	public void constructSetCalendarItemLegacyFreeBusy(){
		CalendarItemType c = new CalendarItemType();
		SetItemFieldType setFreeBusy = factory.constructSetCalendarItemLegacyFreeBusy(c);
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
	
	/**
	 * Test for {@link ExchangeRequestFactory#constructUpdateCalendarItem(CalendarItemType, NonEmptyArrayOfItemChangesType)} 
	 */
	@Test
	public void constructUpdateCalendarItem(){
		CalendarItemType c = new CalendarItemType();
		c.setParentFolderId(createFolderId("calendarFolderId"));
		
		NonEmptyArrayOfItemChangesType changes = new NonEmptyArrayOfItemChangesType();
		ItemChangeType calendarItemChange = new ItemChangeType();
		
		changes.getItemChanges().add(calendarItemChange);
		UpdateItem update = factory.constructUpdateCalendarItem(c, changes);
		assertNotNull(update);
		
		assertNotNull(update.getConflictResolution());
		assertEquals(factory.getConflictResolution(), update.getConflictResolution());
		
		NonEmptyArrayOfItemChangesType itemChanges = update.getItemChanges();
		assertNotNull(itemChanges);
		assertEquals(changes, itemChanges);
		
		//message disposition does not apply to calendars
		assertNull(update.getMessageDisposition());
		
		assertNotNull(update.getSavedItemFolderId());
		assertNotNull(update.getSavedItemFolderId().getFolderId());
		assertEquals(c.getParentFolderId(), update.getSavedItemFolderId().getFolderId());
		
		assertNotNull(update.getSendMeetingInvitationsOrCancellations());
		assertEquals(factory.getSendMeetingUpdates(), update.getSendMeetingInvitationsOrCancellations());
	}

	/**
	 * Test for {@link ExchangeRequestFactory#constructUpdateCalendarItemChanges(CalendarItemType, Collection)}
	 */
	@Test
	public void constructUpdateCalendarItemChanges(){
		CalendarItemType c = new CalendarItemType();
		c.setParentFolderId(createFolderId("calendarFolderId"));
		
		Collection<SetItemFieldType> fields = Collections.singleton(new SetItemFieldType());
		NonEmptyArrayOfItemChangesType changes = factory.constructUpdateCalendarItemChanges(c, fields);
		assertNotNull(changes);
		//TODO
	}

}
