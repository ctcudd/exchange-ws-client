/**
 * 
 */
package com.microsoft.exchange;

import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.xml.bind.JAXBElement;
import javax.xml.datatype.XMLGregorianCalendar;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.Validate;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.util.CollectionUtils;

import com.microsoft.exchange.messages.CreateFolder;
import com.microsoft.exchange.messages.CreateItem;
import com.microsoft.exchange.messages.DeleteFolder;
import com.microsoft.exchange.messages.DeleteItem;
import com.microsoft.exchange.messages.FindFolder;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.messages.GetFolder;
import com.microsoft.exchange.messages.GetItem;
import com.microsoft.exchange.messages.ResolveNames;
import com.microsoft.exchange.messages.UpdateFolder;
import com.microsoft.exchange.types.AffectedTaskOccurrencesType;
import com.microsoft.exchange.types.AndType;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.BaseFolderType;
import com.microsoft.exchange.types.BaseItemIdType;
import com.microsoft.exchange.types.BasePathToElementType;
import com.microsoft.exchange.types.BodyTypeResponseType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemCreateOrDeleteOperationType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.ConstantValueType;
import com.microsoft.exchange.types.ContactItemType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DeleteFolderFieldType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.DistinguishedFolderIdType;
import com.microsoft.exchange.types.ExtendedPropertyType;
import com.microsoft.exchange.types.FieldOrderType;
import com.microsoft.exchange.types.FieldURIOrConstantType;
import com.microsoft.exchange.types.FolderChangeDescriptionType;
import com.microsoft.exchange.types.FolderChangeType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.FolderQueryTraversalType;
import com.microsoft.exchange.types.FolderResponseShapeType;
import com.microsoft.exchange.types.FolderType;
import com.microsoft.exchange.types.IndexBasePointType;
import com.microsoft.exchange.types.IndexedPageViewType;
import com.microsoft.exchange.types.IsGreaterThanOrEqualToType;
import com.microsoft.exchange.types.IsLessThanOrEqualToType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.ItemQueryTraversalType;
import com.microsoft.exchange.types.ItemResponseShapeType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.MessageDispositionType;
import com.microsoft.exchange.types.MessageType;
import com.microsoft.exchange.types.NonEmptyArrayOfAllItemsType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseFolderIdsType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseItemIdsType;
import com.microsoft.exchange.types.NonEmptyArrayOfFieldOrdersType;
import com.microsoft.exchange.types.NonEmptyArrayOfFolderChangeDescriptionsType;
import com.microsoft.exchange.types.NonEmptyArrayOfFolderChangesType;
import com.microsoft.exchange.types.NonEmptyArrayOfFoldersType;
import com.microsoft.exchange.types.NonEmptyArrayOfPathsToElementType;
import com.microsoft.exchange.types.ObjectFactory;
import com.microsoft.exchange.types.PathToExtendedFieldType;
import com.microsoft.exchange.types.PathToUnindexedFieldType;
import com.microsoft.exchange.types.ResolveNamesSearchScopeType;
import com.microsoft.exchange.types.RestrictionType;
import com.microsoft.exchange.types.SetFolderFieldType;
import com.microsoft.exchange.types.SortDirectionType;
import com.microsoft.exchange.types.TargetFolderIdType;
import com.microsoft.exchange.types.TaskType;
import com.microsoft.exchange.types.UnindexedFieldURIType;

/**
 * @author ctcudd
 *
 */
/**
 * @author ctcudd
 *
 */
/**
 * @author ctcudd
 *
 */
public class BaseExchangeRequestFactory {
	

	protected static final int EWS_FIND_ITEM_MAX = 1000;
	
	protected static final int INIT_BASE_OFFSET = 0;
	
	protected final Log log = LogFactory.getLog(this.getClass());
	
    //================================================================================
    // Properties
    //================================================================================
	/**
	 * The default policy in Exchange limits the page size to 1000 items.
	 * Setting the page size to a value that is greater than this number has no
	 * practical effect.
	 * 
	 * @see <a
	 *      href="http://msdn.microsoft.com/en-us/library/office/jj945066(v=exchg.150).aspx#bk_PolicyParameters">EWS
	 *      throttling in Exchange</a>"
	 */
	private int maxFindItems = 500;
	public int getMaxFindItems() {
		return maxFindItems;
	}
	public void setMaxFindItems(int maxFindItems) {
		if(maxFindItems < EWS_FIND_ITEM_MAX){
			this.maxFindItems = maxFindItems;
		}else{
			log.warn("maxFindItems cannot exceed "+EWS_FIND_ITEM_MAX);
		}
	}
	
	/**
	 * The {@link ExtendedPropertyType}'s identified by this {@link Collection} of {@link PathToExtendedFieldType} will be returned by all GetItem operations.
	 */
	private Collection<PathToExtendedFieldType> itemExtendedPropertyPaths = new HashSet<PathToExtendedFieldType>();
	@Autowired(required=false) @Qualifier("extendedPropertiesForItems")
	public void setItemExtendedPropertyPaths(Collection<PathToExtendedFieldType> extendedPropertyPaths) {
		this.itemExtendedPropertyPaths = extendedPropertyPaths;
	}
	public Collection<PathToExtendedFieldType> getItemExtendedPropertyPaths() {
		return itemExtendedPropertyPaths;
	}
	
	/**
	 * The {@link ExtendedPropertyType}'s identified by this  {@link Collection} of {@link PathToExtendedFieldType} will be returned by all GetFolder operations.
	 */
	private Collection<PathToExtendedFieldType> folderExtendedPropertyPaths = new HashSet<PathToExtendedFieldType>();
	/**
	 * @return the folderExtendedPropertyPaths
	 */
	public Collection<PathToExtendedFieldType> getFolderExtendedPropertyPaths() {
		return folderExtendedPropertyPaths;
	}
	/**
	 * @param folderExtendedPropertyPaths the folderExtendedPropertyPaths to set
	 */
	public void setFolderExtendedPropertyPaths(
			Collection<PathToExtendedFieldType> folderExtendedPropertyPaths) {
		this.folderExtendedPropertyPaths = folderExtendedPropertyPaths;
	}

	private final NonEmptyArrayOfPathsToElementType getArrayOfPathsToElementType(Collection<PathToExtendedFieldType> paths){
		NonEmptyArrayOfPathsToElementType arrayOfPaths = new NonEmptyArrayOfPathsToElementType();
		for(PathToExtendedFieldType path : paths){
			arrayOfPaths.getPaths().add(getJaxbPathToExtendedFieldType(path));
		}
		return arrayOfPaths;
	}
	
	private final JAXBElement<PathToExtendedFieldType> getJaxbPathToExtendedFieldType(ExtendedPropertyType extendedPropertyType) {
		PathToExtendedFieldType path = extendedPropertyType.getExtendedFieldURI();
		return getJaxbPathToExtendedFieldType(path);
	}
	
	private final JAXBElement<PathToExtendedFieldType> getJaxbPathToExtendedFieldType(PathToExtendedFieldType path){
		return getObjectFactory().createExtendedFieldURI(path);
	}
	
	private CalendarItemCreateOrDeleteOperationType defaultCalendarCreateSendTo = CalendarItemCreateOrDeleteOperationType.SEND_TO_NONE;
	public CalendarItemCreateOrDeleteOperationType getDefaultCalendarCreateSendTo(){
		return defaultCalendarCreateSendTo;
	}
	
	/**
	 * Returns a new instance of {@link ObjectFactory}
	 * @return {@link ObjectFactory}
	 */
	protected ObjectFactory getObjectFactory() {
		ObjectFactory of = new ObjectFactory();
		return of;
	}
	
    //================================================================================
    // DistinguishedFolderIdType 
    //================================================================================

	/**
	 * @category FolderId
	 * @return {@link DistinguishedFolderIdType} Representing the Primary CALENDAR folder id
	 */
	protected final static DistinguishedFolderIdType getPrimaryCalendarDistinguishedFolderId() {
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.CALENDAR);
	}
	
	/**
	 * @category FolderId
	 * @return {@link DistinguishedFolderIdType} Representing the Primary CONTACTS folder id
	 */
	protected final static DistinguishedFolderIdType getPrimaryContactsDistinguishedFolderId() {
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.CONTACTS);
	}

	/**
	 * @category FolderId
	 * @return {@link DistinguishedFolderIdType} Representing the Primary TASKS folder id
	 */
	protected final static DistinguishedFolderIdType getPrimaryTasksDistinguishedFolderId() {
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.TASKS);
	}

	/**
	 * @category FolderId
	 * @return {@link DistinguishedFolderIdType} Representing the Primary NOTES folder id
	 */
	protected final static DistinguishedFolderIdType getPrimaryNotesDistinguishedFolderId() {
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.NOTES);
	}

	/**
	 * @category FolderId
	 * @return {@link DistinguishedFolderIdType} Representing the Primary NOTES folder id
	 */
	protected final static DistinguishedFolderIdType getPrimaryJournalDistinguishedFolderId() {
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.JOURNAL);
	}
	/**
	 * Construct a {@link DistinguishedFolderIdType} using the {@link DistinguishedFolderIdNameType} as the id value
	 * @category FolderId
	 * @param parent
	 * @return
	 */
	protected final static DistinguishedFolderIdType getParentDistinguishedFolderId( DistinguishedFolderIdNameType parent) {
		DistinguishedFolderIdType distinguishedFolderIdType = new DistinguishedFolderIdType();
		distinguishedFolderIdType.setId(parent);
		return distinguishedFolderIdType;
	}
	
	/**
	 * Construct a {@link TargetFolderIdType} representing the PrimaryFolderId for the given {@link DistinguishedFolderIdNameType}
	 * @category FolderId
	 * @param parent
	 * @return
	 */
	protected final static TargetFolderIdType getParentTargetFolderId( DistinguishedFolderIdNameType parent) {
		TargetFolderIdType targetFolderIdType = new TargetFolderIdType();
		targetFolderIdType.setDistinguishedFolderId(getParentDistinguishedFolderId(parent));
		return targetFolderIdType;
	}
	
	/**
	 * Checks a {@link FolderIdType} for a valid id.  Typically don't care about changeKey
	 * @param folderIdType
	 * @return
	 */
	protected final static boolean isFolderIdValid(FolderIdType folderIdType) {
		boolean isValid = false;
		if(null != folderIdType){
			isValid = StringUtils.isNotBlank(folderIdType.getId());
		}
		return isValid;
	}
	
    //================================================================================
    // ResolveNames
    //================================================================================
	
	/**
	 * Construct a {@link ResolveNames} request
	 * @category ResolveNames
	 * @param alias
	 * @param returnFullContactData
	 * @param searchScope
	 * @param contactDataShape
	 * @return {@link ResolveNames} request
	 */
	protected final ResolveNames constructResolveNames(String alias, boolean returnFullContactData, ResolveNamesSearchScopeType searchScope, DefaultShapeNamesType contactDataShape) {
		ResolveNames resolveNames = new ResolveNames();
		resolveNames.setContactDataShape(contactDataShape);
		resolveNames.setReturnFullContactData(returnFullContactData);
		resolveNames.setSearchScope(searchScope);
		resolveNames.setUnresolvedEntry(alias);
		return resolveNames;
	}
	
    //================================================================================
    // CreateItem
    //================================================================================
	/**
	 * Construct a {@link CreateItem} request
	 * @param list
	 * @param parent
	 * @param dispositionType
	 * @param sendTo
	 * @param folderIdType
	 * @return a {@link CreateItem} request
	 */
	private CreateItem constructCreateItemInternal(Collection<? extends ItemType> list,
			DistinguishedFolderIdNameType parent,
			MessageDispositionType dispositionType,
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderIdType) {

		CreateItem request = new CreateItem();
		NonEmptyArrayOfAllItemsType arrayOfItems = new NonEmptyArrayOfAllItemsType();
		arrayOfItems.getItemsAndMessagesAndCalendarItems().addAll(list);
		request.setItems(arrayOfItems);
		// When the MessageDispositionType is used for the CreateItemType, it only applies to e-mail messages.
		if (null != dispositionType) {
			request.setMessageDisposition(dispositionType);
		}
		TargetFolderIdType tagetFolderId = new TargetFolderIdType();
		if (isFolderIdValid(folderIdType)) {
			// don't set changeKey on createItem as it may have changed in a prior  operation
			FolderIdType fIdType = new FolderIdType();
			fIdType.setId(folderIdType.getId());
			tagetFolderId.setFolderId(fIdType);
		} else if(null != parent ){
			DistinguishedFolderIdType parentDistinguishedFolderId = getParentDistinguishedFolderId(parent);
			tagetFolderId.setDistinguishedFolderId(parentDistinguishedFolderId);
		}else{
			//TODO this should throws?
			log.warn("constructCreateItem cannot determine targetFolderId!");
		}
		request.setSavedItemFolderId(tagetFolderId);
		if (null != sendTo) {
			request.setSendMeetingInvitations(sendTo);
		}
		return request;
	}
	
	/**
	 * Construct a {@link CreateItem} request specific for creating {@link CalendarItemType} objects.
	 * 
	 * This method purposefully omits {@link MessageDispositionType} as it only applies to email messages.
	 * 
	 * If the {@code sendTo} paramater is omitted {@link BaseExchangeRequestFactory#getDefaultCalendarCreateSendTo()} is used.
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa563060(v=exchg.150).aspx">Creating Appointments in Exchange 2010
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx">CreateItem</a>
	 * @param list
	 * @param parent
	 * @param sendTo
	 * @param folderIdType
	 *            - optional. The item will be saved in the Primary calendar
	 *            folder if omitted
	 * @return {@link CreateItem}
	 */
	protected CreateItem constructCreateCalendarItemInternal(Collection<CalendarItemType> list,
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderIdType) {
		if(null == sendTo){
			sendTo = getDefaultCalendarCreateSendTo();
		}
		return constructCreateItemInternal(list, DistinguishedFolderIdNameType.CALENDAR, null, sendTo, folderIdType);
	}
	
	/**
	 * Construct a {@link CreateItem} request specific for creating {@link TaskType} objects
	 * This method purposefully omits {@link MessageDispositionType} as it only applies to email messages.
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx">CreateItem</a>
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa563029(v=exchg.150).aspx">Creating Tasks in Exchange 2010</a>
	 * @param list
	 * @param parent
	 * @param sendTo
	 * @param folderIdType - optional.  The item will be saved in the Primary Tasks folder if omitted
	 * @return {@link CreateItem} 
	 */
	protected CreateItem constructCreateTaskItem(Collection<TaskType> list,
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderIdType) {
		return constructCreateItemInternal(list, DistinguishedFolderIdNameType.TASKS, null, sendTo, folderIdType);
	}
	
	/**
	 * Construct a {@link CreateItem} request specific for creating {@link ContactItemType} objects.
	 * This method purposefully omits {@link MessageDispositionType} as it only applies to email messages.
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx">CreateItem</a>
	 * @see < href="http://msdn.microsoft.com/en-us/library/office/aa563318(v=exchg.150).aspx">Creating Contacts by Using EWS in Exchange 2010</a>
	 * @param items
	 * @param sendTo
	 * @param folderIdType
	 * @return
	 */
	protected CreateItem constructCreateContactItem(Collection<ContactItemType> items, 
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderIdType){
		return constructCreateItemInternal(items, DistinguishedFolderIdNameType.CONTACTS, null, sendTo, folderIdType);
	}
	
	/**
	 * Construct a {@link CreateItem} request specific for creating {@link MessageType} objects.
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx">CreateItem</a>
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa563009(v=exchg.150).aspx">Creating E-mail Messages in Exchange 2010</a>
	 * @param list
	 * @param folderIdType
	 * @return {@link CreateItem} 
	 */
	protected CreateItem constructCreateMessageItemInternal(Collection<MessageType> list, FolderIdType folderIdType) {
		DistinguishedFolderIdNameType parent = DistinguishedFolderIdNameType.INBOX;
		MessageDispositionType disposition = MessageDispositionType.SEND_ONLY;
		CalendarItemCreateOrDeleteOperationType sendTo = CalendarItemCreateOrDeleteOperationType.SEND_ONLY_TO_ALL;
		return constructCreateItemInternal(list, parent, disposition, sendTo, folderIdType);
	}
	
	//================================================================================
    // GetItem
    //================================================================================	
	/**
	 * Construct a {@link GetItem} request for the specified {@link ItemIdType}s
	 * @param itemIds
	 * @param shape
	 * @return {@link GetItem}
	 */
	protected GetItem constructGetItem(Collection<ItemIdType> itemIds, DefaultShapeNamesType shape){
		NonEmptyArrayOfBaseItemIdsType itemIdArray = new NonEmptyArrayOfBaseItemIdsType();
		itemIdArray.getItemIdsAndOccurrenceItemIdsAndRecurringMasterItemIds().addAll(itemIds);
		return constructGetItem(itemIdArray, shape);
	}
	/**
	 * Construct a {@link GetItem} request for the {@link NonEmptyArrayOfBaseItemIdsType}
	 * @param itemIds
	 * @param shape
	 * @return {@link GetItem}
	 */
	protected GetItem constructGetItem(NonEmptyArrayOfBaseItemIdsType items, DefaultShapeNamesType shape){
		ItemResponseShapeType responseShape = constructTextualItemResponseShapeInternal(shape);
		return constructGetItemInternal(items, responseShape);
	}
	
	/**
	 * @param itemIdArray
	 * @param responseShape
	 * @return {@link GetItem}
	 */
	private final GetItem constructGetItemInternal(NonEmptyArrayOfBaseItemIdsType itemIdArray, ItemResponseShapeType responseShape){
		GetItem getItem = new GetItem();
		getItem.setItemIds(itemIdArray);
		getItem.setItemShape(responseShape);
		return getItem;
	}
	
	//================================================================================
    // FindItem
    //================================================================================	
	/**
	 * Constructs a {@link FindItem} request using a {@link CalendarViewType}.  
	 * The maximum number of events returned will be limited by {@link ExchangeRequestFactory#getMaxFindItems()}
	 * A FindItem operation using this request will result in recurrence being expanded by the exchange server, you simply get a list of events (single events & instances of recurring events) as they would appear in the users calendar.  
	 * A FindItem operation using a {@link CalendarViewType} does not support paging, use an {@link IndexedPageViewType} if you need paging.
	 * SearchRestrictions and SortOrders are not valid when using a FindItem operation using a {@link CalendarViewType}
	 * @see http://msdn.microsoft.com/en-us/library/aa564515(v=exchg.140).aspx
	 * @param startTime
	 * @param endTime
	 * @param responseShape
	 * @param traversal
	 * @param folderIds
	 * @return
	 */
	private final FindItem constructCalendarViewFindItem(Date startTime, Date endTime, ItemResponseShapeType responseShape,	ItemQueryTraversalType traversal, NonEmptyArrayOfBaseFolderIdsType arrayOfFolderIds) {
		FindItem findItem = new FindItem();
		findItem.setCalendarView(constructCalendarView(startTime, endTime));
		findItem.setItemShape(responseShape);
		findItem.setTraversal(traversal);
		findItem.setParentFolderIds(arrayOfFolderIds);		
		return findItem;
	}
	
	protected FindItem constructCalendarViewFindItemIdsByDateRange(Date startTime, Date endTime, Collection<FolderIdType> folderIds) {
		NonEmptyArrayOfBaseFolderIdsType arrayOfFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
		arrayOfFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
		ItemResponseShapeType responseShape = constructTextualItemResponseShapeInternal(DefaultShapeNamesType.ID_ONLY);
		return constructCalendarViewFindItem(startTime, endTime, responseShape,  ItemQueryTraversalType.SHALLOW,  arrayOfFolderIds);
	}		
	
	protected FindItem constructIndexedPageViewFindFirstItemIdsShallow(RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		return constructIndexedPageViewFindItemIdsShallow(INIT_BASE_OFFSET, getMaxFindItems(), restriction, folderIds);
	}
	/**
	 * Constructs a {@link FindItem} request using an {@link IndexedPageViewType} (recurrence will not be expanded)
	 * This method performs a Shallow search and will not search sub folders 
	 * This method will return ITEMID only
	 * @param offset
	 * @param maxItems
	 * @param restriction
	 * @param folderIds
	 * @return
	 */
	protected FindItem constructIndexedPageViewFindItemIdsShallow(int offset, int maxItems,  RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		return constructIndexedPageViewFindItem(offset, maxItems, DefaultShapeNamesType.ID_ONLY, ItemQueryTraversalType.SHALLOW, restriction, folderIds);
	}
	
	private FindItem constructIndexedPageViewFindItem(int offset, int maxItems, DefaultShapeNamesType baseShape, ItemQueryTraversalType traversalType, RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		if(maxItems > EWS_FIND_ITEM_MAX){
			log.warn("The default policy in Exchange limits the page size to 1000 items. Setting the page size to a value that is greater than this number has no practical effect. --http://msdn.microsoft.com/en-us/library/office/jj945066(v=exchg.150).aspx#bk_PolicyParameters");
		}
		//use indexed view as restrictions cannot be applied to calendar view
		IndexedPageViewType view = constructIndexedPageView(offset,maxItems,false);
		
		//only return id,  note you can return a limited set of additional properties
		// 	see:http://msdn.microsoft.com/en-us/library/exchange/aa563810(v=exchg.140).aspx
		ItemResponseShapeType responseShape = constructTextualItemResponseShapeInternal(baseShape);
		
		FieldOrderType sortOrder = constructSortOrder();
		List<FieldOrderType> sortOrderList = Collections.singletonList(sortOrder);
		//FindItem findItem = constructIndexedPageViewFindItem(view, responseShape, ItemQueryTraversalType.ASSOCIATED, restriction, sortOrderList, folderIds);
		
		FindItem findItem = constructIndexedPageViewFindItem(view, responseShape, traversalType, restriction, sortOrderList, folderIds);
		return findItem;
	}
	
	/**
	 * 
	 * @param view
	 * @param responseShape
	 * @param traversal
	 * @param restriction
	 * @param sortOrderList
	 * @param folderIds
	 * @return
	 */
	private FindItem constructIndexedPageViewFindItem(
			IndexedPageViewType view, ItemResponseShapeType responseShape,
			ItemQueryTraversalType traversal, RestrictionType restriction,
			Collection<FieldOrderType> sortOrderList,
			Collection<? extends BaseFolderIdType> folderIds) {
		FindItem findItem = new FindItem();

		findItem.setIndexedPageItemView(view);     
		findItem.setItemShape(responseShape);
		findItem.setTraversal(traversal);

		if (null != restriction) {
			findItem.setRestriction(restriction);
		}

		if (!CollectionUtils.isEmpty(sortOrderList)) {
			NonEmptyArrayOfFieldOrdersType sortOrder = new NonEmptyArrayOfFieldOrdersType();
			sortOrder.getFieldOrders().addAll(sortOrderList);
			findItem.setSortOrder(sortOrder);
		}

		if (!CollectionUtils.isEmpty(folderIds)) {
			NonEmptyArrayOfBaseFolderIdsType parentFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
			parentFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(
					folderIds);
			findItem.setParentFolderIds(parentFolderIds);
		}

		return findItem;
	}
	
	//================================================================================
    // DeleteItem
    //================================================================================	
	private final  DeleteItem constructDeleteItemInternal(
			NonEmptyArrayOfBaseItemIdsType arrayOfItemIds,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendTo,
			AffectedTaskOccurrencesType affectedTaskOccurrencesType) {
		DeleteItem deleteItem = new DeleteItem();
		deleteItem.setDeleteType(disposalType);
		deleteItem.setItemIds(arrayOfItemIds);
		if (null != affectedTaskOccurrencesType) {
			deleteItem.setAffectedTaskOccurrences(affectedTaskOccurrencesType);
		}
		if (null != sendTo) {
			deleteItem.setSendMeetingCancellations(sendTo);
		}
		return deleteItem;
	}
	
	protected DeleteItem constructDeleteItem(
			Collection<? extends BaseItemIdType> itemIds,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendTo,
			AffectedTaskOccurrencesType affectedTaskOccurrencesType) {
		Validate.notEmpty(itemIds, "must specify at least one itemId.");
		Validate.notNull(disposalType, "disposalType cannot be null");
		NonEmptyArrayOfBaseItemIdsType arrayOfItemIds = new NonEmptyArrayOfBaseItemIdsType();
		arrayOfItemIds
				.getItemIdsAndOccurrenceItemIdsAndRecurringMasterItemIds()
				.addAll(itemIds);
		return constructDeleteItemInternal(arrayOfItemIds, disposalType, sendTo, affectedTaskOccurrencesType);
	}
	
    //================================================================================
    // CreateFolder - methods to aid in constucting CreateFolder requests
    //================================================================================
	/**
	 * Construct a {@link CreateFolder} request, specifying the {@code parent} as the parent.
	 * @param parent - where the folders are to be created
	 * @param folders - the folders to create
	 * @return {@link CreateFolder}
	 */
	protected CreateFolder constructCreateFolder(
			DistinguishedFolderIdNameType parent,
			Collection<? extends BaseFolderType> folders) {
		return constructCreateFolderInternal(getParentTargetFolderId(parent), folders);
	}
	
	/**
	 * Construct a {@link CreateFolder} request.
	 * 
	 * This method is private because not all {@link TargetFolderIdType}s are valid.
	 * Example: You can only create {@link CalendarFolderType} within the {@link DistinguishedFolderIdNameType#CALENDAR}.
	 * 			You cannot create {@link CalendarFolderType} within a {@link FolderIdType}, unless the {@link FolderIdType} resolves to the {@link DistinguishedFolderIdNameType#CALENDAR}
	 * 
	 * @param parent - where the folders are to be created
	 * @param folders - the folders to create
	 * @return {@link CreateFolder}
	 */
	private CreateFolder constructCreateFolderInternal(TargetFolderIdType parent, Collection<? extends BaseFolderType> folders){
		CreateFolder createFolder = new CreateFolder();
		createFolder.setParentFolderId(parent);
		NonEmptyArrayOfFoldersType folderArray = new NonEmptyArrayOfFoldersType();
		folderArray.getFoldersAndCalendarFoldersAndContactsFolders().addAll(folders);
		createFolder.setFolders(folderArray);
		return createFolder;
	}
	
	//================================================================================
    // GetFolder
    //================================================================================	
	/**
	 * Construct a {@link GetFolder} request
	 * @param baseFolderIds
	 * @param responseShape
	 * @return {@link GetFolder}
	 */
	private final GetFolder constructGetFolderInternal(NonEmptyArrayOfBaseFolderIdsType baseFolderIds, FolderResponseShapeType responseShape){
		GetFolder getFolder = new GetFolder();
		getFolder.setFolderIds(baseFolderIds);
		getFolder.setFolderShape(responseShape);
		return getFolder;
	}
	
	/**
	 * Construct a {@link FolderResponseShapeType}
	 * 
	 * @param shape
	 * @param exProps
	 * @return {@link FolderResponseShapeType}
	 */
	private final FolderResponseShapeType constructFolderResponseShapeInternal(DefaultShapeNamesType shape, NonEmptyArrayOfPathsToElementType exProps){
		FolderResponseShapeType responseShape = new FolderResponseShapeType();
		responseShape.setAdditionalProperties(exProps);
		responseShape.setBaseShape(shape);
		return responseShape;
	}
	
	/**
	 * Construct a {@link GetFolder} request which will retrieve {@link DefaultShapeNamesType#ALL_PROPERTIES} for the given {@link FolderIdType}s
	 *
	 * @param folderIdArray
	 * @return {@link GetFolder}
	 */
	protected GetFolder constructGetFolder(NonEmptyArrayOfBaseFolderIdsType folderIdArray){
		NonEmptyArrayOfPathsToElementType paths = getArrayOfPathsToElementType(getFolderExtendedPropertyPaths());
		FolderResponseShapeType responseShape = constructFolderResponseShapeInternal(DefaultShapeNamesType.ALL_PROPERTIES, paths);
		return constructGetFolderInternal(folderIdArray, responseShape);
	}
	
	//================================================================================
    // FindFolder
    //================================================================================	
	protected final FindFolder constructFindFolderInternal(DistinguishedFolderIdNameType parent,
			DefaultShapeNamesType folderShape,
			FolderQueryTraversalType folderQueryTraversalType,
			RestrictionType restriction){
		
		FindFolder findFolder = new FindFolder();
		findFolder.setTraversal(folderQueryTraversalType);

		FolderResponseShapeType responseShape = new FolderResponseShapeType();
		responseShape.setBaseShape(folderShape);
		findFolder.setFolderShape(responseShape);

		IndexedPageViewType pageView = constructIndexedPageView(INIT_BASE_OFFSET, EWS_FIND_ITEM_MAX, false);
		findFolder.setIndexedPageFolderView(pageView);

		DistinguishedFolderIdType parentDistinguishedFolderId = getParentDistinguishedFolderId(parent);
		NonEmptyArrayOfBaseFolderIdsType array = new NonEmptyArrayOfBaseFolderIdsType();
		array.getFolderIdsAndDistinguishedFolderIds().add(parentDistinguishedFolderId);
		findFolder.setParentFolderIds(array);

		if (null != restriction) {
			findFolder.setRestriction(restriction);
		}
		return findFolder;
	}
	
	//================================================================================
    // UpdateFolder
    //================================================================================	
	protected UpdateFolder constructUpdateFolderDeleteExtendedProperty(
			FolderIdType folderId, ExtendedPropertyType exProp) {
		return constructUpdateFolderDeleteField(
				getJaxbPathToExtendedFieldType(exProp), folderId);
	}

	protected UpdateFolder constructUpdateFolderSetField(FolderType folder,
			JAXBElement<? extends BasePathToElementType> path,
			FolderIdType folderId) {
		SetFolderFieldType changeDescription = new SetFolderFieldType();
		changeDescription.setFolder(folder);
		changeDescription.setPath(path);
		return constructUpdateFolderInternal(changeDescription, folderId);
	}

	protected UpdateFolder constructUpdateFolderDeleteField(
			JAXBElement<? extends BasePathToElementType> path,
			FolderIdType folderId) {
		DeleteFolderFieldType changeDescription = new DeleteFolderFieldType();
		changeDescription.setPath(path);
		return constructUpdateFolderInternal(changeDescription, folderId);
	}

	private UpdateFolder constructUpdateFolderInternal(
			FolderChangeDescriptionType changeDescription, FolderIdType folderId) {

		NonEmptyArrayOfFolderChangeDescriptionsType folderUpdates = new NonEmptyArrayOfFolderChangeDescriptionsType();
		folderUpdates
				.getAppendToFolderFieldsAndSetFolderFieldsAndDeleteFolderFields()
				.add(changeDescription);

		FolderChangeType folderChange = new FolderChangeType();
		folderChange.setFolderId(folderId);
		folderChange.setUpdates(folderUpdates);

		NonEmptyArrayOfFolderChangesType changes = new NonEmptyArrayOfFolderChangesType();
		changes.getFolderChanges().add(folderChange);

		UpdateFolder updateRequest = new UpdateFolder();
		updateRequest.setFolderChanges(changes);

		return updateRequest;
	}
	
	//================================================================================
    // DeleteFolder
    //================================================================================
	/**
	 * Constructs a {@link DeleteFolder} request
	 * @param folderIds
	 * @param disposalType
	 * @return {@link DeleteFolder}
	 */
	protected final DeleteFolder constructDeleteFolderInternal(
			NonEmptyArrayOfBaseFolderIdsType folderIdArray,
			DisposalType disposalType) {
		DeleteFolder deleteFolder = new DeleteFolder();
		deleteFolder.setDeleteType(disposalType);
		deleteFolder.setFolderIds(folderIdArray);
		return deleteFolder;
	}
	/**
	 * Constructs a {@link DeleteFolder} request
	 * @param folderIds
	 * @param disposalType
	 * @return {@link DeleteFolder}
	 */
	protected DeleteFolder constructDeleteFolder(
			Collection<? extends BaseFolderIdType> folderIds,
			DisposalType disposalType) {
		Validate.notEmpty(folderIds, "folderIds cannot be empty");
		NonEmptyArrayOfBaseFolderIdsType folderIdArray = new NonEmptyArrayOfBaseFolderIdsType();
		folderIdArray.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
		return constructDeleteFolderInternal(folderIdArray, disposalType);
	}
	
	//================================================================================
    // CalendarViewType
    //================================================================================
	/**
	 * The {@link CalendarViewType} element defines a FindItem Operation as returning calendar items in a set as they appear in a calendar.
	 * 
	 * If the CalendarView element is specified in a FindItem request, the Web service returns a list of single calendar items and occurrences of recurring calendar items within the range specified by StartDate and EndDate.
	 * 
	 * see: http://msdn.microsoft.com/en-us/library/aa564515(v=exchg.140).aspx
	 * 
	 * @param startTime
	 * @param endTime
	 * @return {@link CalendarViewType}
	 */
	protected CalendarViewType constructCalendarView(Date startTime, Date endTime) {
		CalendarViewType calendarView = new CalendarViewType();
		calendarView.setMaxEntriesReturned(getMaxFindItems());
		calendarView.setStartDate(ExchangeDateUtils
				.convertDateToXMLGregorianCalendar(startTime));
		calendarView.setEndDate(ExchangeDateUtils
				.convertDateToXMLGregorianCalendar(endTime));

		return calendarView;
	}
	
	//================================================================================
    // IndexedPageViewType
    //================================================================================
	/**
	 * Constructs an {@link IndexedPageViewType}
	 * @param start
	 * @param length
	 * @param reverse
	 * @return {@link IndexedPageViewType}
	 */
	private IndexedPageViewType constructIndexedPageView(Integer start,
			Integer length, Boolean reverse) {
		IndexedPageViewType view = new IndexedPageViewType();
		view.setMaxEntriesReturned(length);
		view.setOffset(start);
		if (reverse) {
			view.setBasePoint(IndexBasePointType.END);
		} else {
			view.setBasePoint(IndexBasePointType.BEGINNING);
		}
		return view;
	}
	
    //================================================================================
    // Restrictions - methods to aid in constucting Search Restrictions
	// (Typically used with Find<Item|Folder> operations)
    //================================================================================
	/**
	 * Construct a {@link RestrictionType} object for use with a {@link FindItem} using a {@link CalendarViewType}.  
	 * @param startTime
	 * @param endTime
	 * @return a {@link RestrictionType} which targets {@link CalendarItemType} betweem {@code startTime} and {@code endTime}
	 */
	protected RestrictionType constructFindCalendarItemsByDateRangeRestriction(Date startTime, Date endTime) {
		ObjectFactory of = getObjectFactory();
		JAXBElement<IsGreaterThanOrEqualToType> startSearchExpression = getCalendarItemStartSearchExpression(startTime);
		
		JAXBElement<IsLessThanOrEqualToType> endSearchExpression = getCalendarItemEndSearchExpression( endTime);
		
		//and them all together
		AndType andType = new AndType();
		andType.getSearchExpressions().add(startSearchExpression);
		andType.getSearchExpressions().add(endSearchExpression);
		JAXBElement<AndType> andSearchExpression = of.createAnd(andType);
		
		// create restriction and set searchExpression
		RestrictionType restrictionType = new RestrictionType();
		restrictionType.setSearchExpression(andSearchExpression);
		return restrictionType;
	}
	
	/**
	 * Construct a search expression for finding calendar itmems ON or BEFORE the specified {@code endTime}
	 * @param endTime
	 * @return {@link JAXBElement<IsLessThanOrEqualToType>}
	 */
	private final JAXBElement<IsLessThanOrEqualToType> getCalendarItemEndSearchExpression(Date endTime) {
		ObjectFactory of = getObjectFactory();
		IsLessThanOrEqualToType endType = new IsLessThanOrEqualToType();
		XMLGregorianCalendar end = ExchangeDateUtils.convertDateToXMLGregorianCalendar(endTime);
		PathToUnindexedFieldType endPath = new PathToUnindexedFieldType();
		endPath.setFieldURI(UnindexedFieldURIType.CALENDAR_END);
		JAXBElement<PathToUnindexedFieldType> endFieldURI = of.createFieldURI(endPath);
		endType.setPath(endFieldURI);
		FieldURIOrConstantType endConstant = new FieldURIOrConstantType();
		ConstantValueType endValue = new ConstantValueType();
		endValue.setValue(end.toXMLFormat());
		endConstant.setConstant(endValue);
		endType.setFieldURIOrConstant(endConstant);
		JAXBElement<IsLessThanOrEqualToType> endSearchExpression = of.createIsLessThanOrEqualTo(endType);
		return endSearchExpression;
	}
	
	/**
	 * Construct a search expression for finding calendar itmems ON or AFTER the specified {@code startTime}
	 * @param endTime
	 * @return {@link JAXBElement<IsGreaterThanOrEqualToType>}
	 */
	private final JAXBElement<IsGreaterThanOrEqualToType> getCalendarItemStartSearchExpression(Date startTime) {
		ObjectFactory of = getObjectFactory();
		IsGreaterThanOrEqualToType startType = new IsGreaterThanOrEqualToType();
		XMLGregorianCalendar start = ExchangeDateUtils.convertDateToXMLGregorianCalendar(startTime);
		PathToUnindexedFieldType startPath = new PathToUnindexedFieldType();
		startPath.setFieldURI(UnindexedFieldURIType.CALENDAR_START);
		JAXBElement<PathToUnindexedFieldType> startFieldURI = of.createFieldURI(startPath);
		startType.setPath(startFieldURI);
		FieldURIOrConstantType startConstant = new FieldURIOrConstantType();
		ConstantValueType startValue = new ConstantValueType();
		startValue.setValue(start.toXMLFormat());
		startConstant.setConstant(startValue);
		startType.setFieldURIOrConstant(startConstant);
		JAXBElement<IsGreaterThanOrEqualToType> startSearchExpression = of.createIsGreaterThanOrEqualTo(startType);
		return startSearchExpression;
	}
	
	//================================================================================
    // ItemResponseShapeType
    //================================================================================	
	/**
	 * Constructs an {@link ItemResponseShapeType} specifying {@link BodyTypeResponseType#TEXT} as the {@link BodyTypeResponseType}
	 * @param baseShape
	 * @return
	 */
	private final ItemResponseShapeType constructTextualItemResponseShapeInternal(DefaultShapeNamesType baseShape) {
		return constructItemResponseShapeInternal(baseShape, BodyTypeResponseType.TEXT, true, true, false);
	}
	/**
	 * Constructs an {@link ItemResponseShapeType} 
	 * This methods uses {@link BaseExchangeRequestFactory#getItemExtendedPropertyPaths()} as a source for additional properties 
	 * @param baseShape
	 * @param bodyType
	 * @param htmlToUtf8
	 * @param filterHtml
	 * @param includeMime
	 * @return
	 */
	private final ItemResponseShapeType constructItemResponseShapeInternal(
			DefaultShapeNamesType baseShape, BodyTypeResponseType bodyType,
			Boolean htmlToUtf8, Boolean filterHtml, Boolean includeMime) {

		ItemResponseShapeType responseShape = new ItemResponseShapeType();
		responseShape.setBaseShape(baseShape);
		if (null != bodyType) {
			responseShape.setBodyType(bodyType);
		}
		if (null != htmlToUtf8) {
			responseShape.setConvertHtmlCodePageToUTF8(htmlToUtf8);
		}
		if (null != filterHtml) {
			responseShape.setFilterHtmlContent(filterHtml);
		}
		if (null != includeMime) {
			responseShape.setIncludeMimeContent(includeMime);
		}
		Collection<PathToExtendedFieldType> extendedPropertyPaths = getItemExtendedPropertyPaths();
		if (!CollectionUtils.isEmpty(extendedPropertyPaths) ) {
			responseShape.setAdditionalProperties(getArrayOfPathsToElementType(extendedPropertyPaths));
		}
		return responseShape;
	}
	
	//================================================================================
    // FieldOrderType
    //================================================================================		
	/**
	 * Constructs a {@link FieldOrderType} which will sort items by id ascending.
	 * 
	 * @return
	 */
	protected FieldOrderType constructSortOrder() {
		ObjectFactory of = getObjectFactory();
		//set sort order (earliest items first)
		FieldOrderType sortOrder = new FieldOrderType();
		sortOrder.setOrder(SortDirectionType.ASCENDING);
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		path.setFieldURI(UnindexedFieldURIType.ITEM_ITEM_ID);
		JAXBElement<PathToUnindexedFieldType> sortPath = of.createFieldURI(path);
		sortOrder.setPath(sortPath);
		return sortOrder;
	}
}
