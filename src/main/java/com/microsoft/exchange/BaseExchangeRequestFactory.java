/**
 * 
 */
package com.microsoft.exchange;

import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import javax.xml.bind.JAXBElement;
import javax.xml.datatype.XMLGregorianCalendar;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.Validate;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
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
import com.microsoft.exchange.messages.ResolveNames;
import com.microsoft.exchange.messages.UpdateFolder;
import com.microsoft.exchange.messages.UpdateItem;
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
import com.microsoft.exchange.types.CalendarItemUpdateOperationType;
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.ConflictResolutionType;
import com.microsoft.exchange.types.ConstantValueType;
import com.microsoft.exchange.types.ContactItemType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.DeleteFolderFieldType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.DistinguishedFolderIdNameType;
import com.microsoft.exchange.types.DistinguishedFolderIdType;
import com.microsoft.exchange.types.EmailAddressType;
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
import com.microsoft.exchange.types.IsLessThanType;
import com.microsoft.exchange.types.ItemChangeType;
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
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangeDescriptionsType;
import com.microsoft.exchange.types.NonEmptyArrayOfItemChangesType;
import com.microsoft.exchange.types.NonEmptyArrayOfPathsToElementType;
import com.microsoft.exchange.types.ObjectFactory;
import com.microsoft.exchange.types.PathToExtendedFieldType;
import com.microsoft.exchange.types.PathToUnindexedFieldType;
import com.microsoft.exchange.types.ResolveNamesSearchScopeType;
import com.microsoft.exchange.types.RestrictionType;
import com.microsoft.exchange.types.SearchExpressionType;
import com.microsoft.exchange.types.SetFolderFieldType;
import com.microsoft.exchange.types.SetItemFieldType;
import com.microsoft.exchange.types.TargetFolderIdType;
import com.microsoft.exchange.types.TaskType;
import com.microsoft.exchange.types.UnindexedFieldURIType;

/**
 * 
 * @author ctcudd
 */
public abstract class BaseExchangeRequestFactory {
	
    //================================================================================
    // Static Fields
    //================================================================================
	protected static final int FIND_ITEM_MAX 	= 1000;
	protected static final int INIT_BASE_OFFSET = 0;
	public static final String SMTP = "smtp:";
	
    //================================================================================
    // Properties
    //================================================================================
	protected final Log log = LogFactory.getLog(this.getClass());
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
	
    //================================================================================
    // Getters
    //================================================================================
	/**
	 * @return the {@link AffectedTaskOccurrencesType} to use with operations on {@link TaskType}s.
	 */
	public abstract AffectedTaskOccurrencesType getAffectedTaskOccurrencesType();
	/**
	 * @return the {@link ConflictResolutionType} to use when updating items.
	 */
	public abstract ConflictResolutionType getConflictResolution();
	/**
	 * @return the {@link FolderQueryTraversalType}
	 */
	public abstract FolderQueryTraversalType getFolderQueryTraversal();
	
	/**
	 * @return the {@link ItemQueryTraversalType}.
	 */
	public abstract ItemQueryTraversalType getItemQueryTraversal();
	
	/**
	 * The {@link ExtendedPropertyType}'s identified by this {@link Collection} of {@link PathToExtendedFieldType} will be returned by all GetItem operations.
	 */
	public abstract Collection<PathToExtendedFieldType> getItemExtendedPropertyPaths();
	/**
	 * The {@link ExtendedPropertyType}s identified by this  {@link Collection} of {@link PathToExtendedFieldType} will be returned by all GetFolder operations.
	 */
	public abstract Collection<PathToExtendedFieldType> getFolderExtendedPropertyPaths();
	/**
	 * @return the {@link CalendarItemCreateOrDeleteOperationType} to use when performing operations on {@link CalendarItemType}s.
	 */
	public abstract CalendarItemCreateOrDeleteOperationType getSendMeetingInvitations();
	/**
	 * @return
	 */
	public abstract CalendarItemUpdateOperationType getSendMeetingUpdates();
	/**
	 * @return the maximum number of items that will be retrieved in a {@link FindItem} operation.
	 */
	public final int getMaxFindItems() {
		if(maxFindItems > FIND_ITEM_MAX){
			maxFindItems = FIND_ITEM_MAX;
			log.warn("maxFindItems cannot exceed "+FIND_ITEM_MAX);
		}
		return maxFindItems;
	}
	/**
	 * @return the {@link MessageDispositionType} to use when creating {@link MessageType}s.  This does not apply to {@link CalendarItemType} or {@link TaskType}.
	 */
	public abstract MessageDispositionType getMessageDisposition();
	
	public abstract ResolveNamesSearchScopeType getResolveNamesSearchScope();
	
	//================================================================================
    // Setters
    //================================================================================	
	/**
	 * @param maxFindItems the maxFindItems to set.
	 */
	public final void setMaxFindItems(int maxFindItems) {
		if(maxFindItems < FIND_ITEM_MAX){
			this.maxFindItems = maxFindItems;
		}else{
			this.maxFindItems = FIND_ITEM_MAX;
			log.warn("maxFindItems cannot exceed "+FIND_ITEM_MAX);
		}
	}
    //================================================================================
    // MEAT
    //================================================================================	
	/**
	 * Construct a {@link NonEmptyArrayOfPathsToElementType} given a collection of {@link PathToExtendedFieldType}s
	 * @param paths
	 * @return
	 */
	private final NonEmptyArrayOfPathsToElementType getArrayOfPathsToElementType(Collection<PathToExtendedFieldType> paths){
		NonEmptyArrayOfPathsToElementType arrayOfPaths = new NonEmptyArrayOfPathsToElementType();
		for(PathToExtendedFieldType path : paths){
			arrayOfPaths.getPaths().add(getJaxbPathToExtendedFieldType(path));
		}
		return arrayOfPaths;
	}
	
	/**
	 * Construct a {@link JAXBElement<PathToExtendedFieldType>} given a {@link ExtendedPropertyType}
	 * @param extendedPropertyType
	 * @return
	 */
	private final JAXBElement<PathToExtendedFieldType> getJaxbPathToExtendedFieldType(ExtendedPropertyType extendedPropertyType) {
		PathToExtendedFieldType path = extendedPropertyType.getExtendedFieldURI();
		return getJaxbPathToExtendedFieldType(path);
	}
	
	/**
	 * Construct a  {@link JAXBElement<PathToExtendedFieldType>} given a {@link PathToExtendedFieldType}
	 * @param path
	 * @return
	 */
	private final JAXBElement<PathToExtendedFieldType> getJaxbPathToExtendedFieldType(PathToExtendedFieldType path){
		return getObjectFactory().createExtendedFieldURI(path);
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
	 * @return {@link DistinguishedFolderIdType} Representing the Primary Message folder id (aka the INBOX)
	 */
	protected final static DistinguishedFolderIdType getPrimaryMessageDistinguishedFolderId(){
		return getParentDistinguishedFolderId(DistinguishedFolderIdNameType.INBOX);
	}
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
	private final ResolveNames constructResolveNames(String alias, boolean returnFullContactData, ResolveNamesSearchScopeType searchScope, DefaultShapeNamesType contactDataShape) {
		ResolveNames resolveNames = new ResolveNames();
		resolveNames.setContactDataShape(contactDataShape);
		resolveNames.setReturnFullContactData(returnFullContactData);
		resolveNames.setSearchScope(searchScope);
		resolveNames.setUnresolvedEntry(alias);
		return resolveNames;
	}
	
	/**
	 * When using impersonation, your admin account should have a mailbox, if it doesn't you may run into an error like the following:
	 * ERROR_MISSING_EMAIL_ADDRESS - "When making a request as an account that does not have a mailbox, you must specify the mailbox primary SMTP address for any distinguished folder Ids.".
	 * 
	 * Unlike {@link #constructResolveNames(String, boolean, ResolveNamesSearchScopeType, DefaultShapeNamesType)}, this method will the set the parentFolderIds property of the resulting {@link ResolveNames} object
	 * 
	 * @param alias
	 * @param setDistinqishedFolderId
	 * @param returnFullContactData
	 * @param searchScope
	 * @param contactDataShape
	 * @return
	 */
	protected ResolveNames constructResolveNames(String alias, boolean setDistinqishedFolderId, boolean returnFullContactData, ResolveNamesSearchScopeType searchScope, DefaultShapeNamesType contactDataShape) {
		String emailAddress = alias;
		if(alias.startsWith(SMTP)){
			emailAddress = alias.substring(SMTP.length());
		}else{
			alias = SMTP + alias;
		}
		ResolveNames request = constructResolveNames(alias, returnFullContactData, searchScope, contactDataShape);
		if(setDistinqishedFolderId){
			NonEmptyArrayOfBaseFolderIdsType baseFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
			DistinguishedFolderIdType baseFolderId = new DistinguishedFolderIdType();
			baseFolderId.setId(DistinguishedFolderIdNameType.CONTACTS);
			EmailAddressType mailbox = new EmailAddressType();
			mailbox.setEmailAddress(emailAddress);
			baseFolderId.setMailbox(mailbox);
			baseFolderIds.getFolderIdsAndDistinguishedFolderIds().add(baseFolderId);
			request.setParentFolderIds(baseFolderIds);
		}
		return request;
	}
    //================================================================================
    // CreateItem
    //================================================================================
	/**
	 * Construct a {@link CreateItem} request specific for creating {@link CalendarItemType} objects.
	 * 
	 * This method purposefully omits {@link MessageDispositionType} as it only applies to email messages.
	 * 
	 * If the {@code sendTo} paramater is omitted {@link BaseExchangeRequestFactory#getDefaultCalendarCreateSendTo()} is used.
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa563060(v=exchg.150).aspx">Creating Appointments in Exchange 2010
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa565209(v=exchg.150).aspx">CreateItem</a>
	 * @category CreateItem CalendarItem
	 * 
	 * @param calendarItems
	 * @param sendTo
	 * @param folderId
	 * @return {@link CreateItem}
	 */
	protected final CreateItem constructCreateCalendarItem(
			Collection<CalendarItemType> calendarItems,
			CalendarItemCreateOrDeleteOperationType sendTo,
			FolderIdType folderId) {
		return constructCreateItemInternal(calendarItems, DistinguishedFolderIdNameType.CALENDAR, null, sendTo, folderId);
	}
	/**
	 * Construct a {@link CreateItem} request
	 * @param list
	 * @param parent
	 * @param dispositionType
	 * @param sendTo
	 * @param folderIdType
	 * @return a {@link CreateItem} request
	 */
	private final CreateItem constructCreateItemInternal(Collection<? extends ItemType> list,
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
	protected final CreateItem constructCreateTaskItem(Collection<TaskType> list,
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
		return constructCreateItemInternal(list, DistinguishedFolderIdNameType.INBOX, getMessageDisposition(), null, folderIdType);
	}

	//================================================================================
    // UpdateItem
    //================================================================================	
	
	/**
	 * Construct an {@link UpdateItem} request which when executed will apply the {@code changes} to the {@code calendarItem}
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa580254(v=exchg.150).aspx">UpdateItem</a>
	 * @see <a href="http://msdn.microsoft.com/en-us/library/office/aa581084(v=exchg.150).aspx>UpdateItem operation</a>
	 * @param calendarItem
	 * @param changes
	 * @return
	 */
	public UpdateItem constructUpdateCalendarItem(CalendarItemType calendarItem, NonEmptyArrayOfItemChangesType changes){
		UpdateItem updateItem = new UpdateItem();
		updateItem.setConflictResolution(getConflictResolution());
		updateItem.setItemChanges(changes);
		
		//messageDisposition required only for message types
		//updateItem.setMessageDisposition(MessageDispositionType.);
		
		TargetFolderIdType targetFolderId = new TargetFolderIdType();
		targetFolderId.setFolderId(calendarItem.getParentFolderId());
		
		// TODO the update will fail if  null == calendarItem.getParentFolderId().  Consider throwing / warning here.
		
		updateItem.setSavedItemFolderId(targetFolderId);
		updateItem.setSendMeetingInvitationsOrCancellations(getSendMeetingUpdates());
		return updateItem;
	}
	
	/**
	 * Construct a {@link NonEmptyArrayOfItemChangesType} for a give {@link CalendarItemType} and a {@link Collection} of {@link SetItemFieldType}s
	 * @param item
	 * @param fields
	 * @return
	 */
	public NonEmptyArrayOfItemChangesType constructUpdateCalendarItemChanges(CalendarItemType item, Collection<SetItemFieldType> fields){
		NonEmptyArrayOfItemChangesType changes = new NonEmptyArrayOfItemChangesType();
		NonEmptyArrayOfItemChangeDescriptionsType updates = new NonEmptyArrayOfItemChangeDescriptionsType();
		for(SetItemFieldType field : fields){
			updates.getAppendToItemFieldsAndSetItemFieldsAndDeleteItemFields().add(field);
		}
		ItemChangeType change = new ItemChangeType();
		change.setItemId(item.getItemId());
		change.setUpdates(updates);
		changes.getItemChanges().add(change);
		return changes;
	}
	/**
	 * Construct a {@link SetItemFieldType} which can be used to update a {@link CalendarItemType}s subject field
	 * @param item
	 * @return
	 */
	public SetItemFieldType constructSetCalendarItemSubject(CalendarItemType item){
		CalendarItemType c = new CalendarItemType();
		c.setSubject(item.getSubject());
		SetItemFieldType changeDescription = new SetItemFieldType();
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		path.setFieldURI(UnindexedFieldURIType.ITEM_SUBJECT);
		changeDescription.setPath(getObjectFactory().createPath(path));
		changeDescription.setCalendarItem(c);
		return changeDescription;
	}
	
	/**
	 * Construct a {@link SetItemFieldType} which can be used to update a {@link CalendarItemType}s LegacyFreeBusy field
	 * @param item
	 * @return
	 */
	public SetItemFieldType constructSetCalendarItemLegacyFreeBusy(CalendarItemType item){
		CalendarItemType c = new CalendarItemType();
		c.setLegacyFreeBusyStatus(item.getLegacyFreeBusyStatus());
		SetItemFieldType changeDescription = new SetItemFieldType();
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		path.setFieldURI(UnindexedFieldURIType.CALENDAR_LEGACY_FREE_BUSY_STATUS);
		changeDescription.setPath(getObjectFactory().createPath(path));
		changeDescription.setCalendarItem(c);
		return changeDescription;
	}
	
	/**
	 * Construct a {@link SetItemFieldType} which can be used to update a {@link CalendarItemType}s RequiredAttendees field
	 * @param item
	 * @return
	 */
	public SetItemFieldType constructSetCalendarItemRequiredAttendees(CalendarItemType item){
		CalendarItemType c = new CalendarItemType();
		c.setRequiredAttendees(item.getRequiredAttendees());
		SetItemFieldType changeDescription = new SetItemFieldType();
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		path.setFieldURI(UnindexedFieldURIType.CALENDAR_REQUIRED_ATTENDEES);
		changeDescription.setPath(getObjectFactory().createPath(path));
		changeDescription.setCalendarItem(c);
		return changeDescription;
	}
	
	/**
	 * Construct {@link SetItemFieldType} which can be used to update a {@link CalendarItemType}s {@link ExtendedPropertyType}
	 * @param item
	 * @param exprop
	 * @return
	 */
	public SetItemFieldType constructSetCalendarItemExtendedProperty(CalendarItemType item, ExtendedPropertyType exprop){
		CalendarItemType c = new CalendarItemType();
		c.getExtendedProperties().add(exprop);
		SetItemFieldType changeDescription = new SetItemFieldType();
		PathToExtendedFieldType extendedFieldURI = exprop.getExtendedFieldURI();
		changeDescription.setPath(getObjectFactory().createPath(extendedFieldURI));
		changeDescription.setCalendarItem(c);
		return changeDescription;
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
	 * A FindItem operation using this request will result in recurrence being expanded by the exchange server, you simply get a list of events (single events & instances of recurring events, NO recurring masters) as they would appear in the users calendar.  
	 * A FindItem operation using a {@link CalendarViewType} does not support paging, use an {@link IndexedPageViewType} if you need paging/sorting.
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
	
	/**
	 * Construct a {@link FindItem} request which will perform a {@link ItemQueryTraversalType#SHALLOW} (non-recursive) search of the given {@link BaseFolderIdType}s.
	 * This method uses {@link CalendarViewType}, which does not support paging/sorting but expands recurrence for you.
	 * @param startTime
	 * @param endTime
	 * @param folderIds
	 * @return
	 */
	protected FindItem constructCalendarViewFindItemIdsByDateRange(Date startTime, Date endTime, Collection<? extends BaseFolderIdType> folderIds) {
		NonEmptyArrayOfBaseFolderIdsType arrayOfFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
		arrayOfFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
		ItemResponseShapeType responseShape = constructTextualItemResponseShapeInternal(DefaultShapeNamesType.ID_ONLY);
		return constructCalendarViewFindItem(startTime, endTime, responseShape,  ItemQueryTraversalType.SHALLOW,  arrayOfFolderIds);
	}		
	
	/**
	 * Construct a {@link FindItem} rquest which will search the specified {@link BaseFolderIdType}s using the specified {@link RestrictionType}.
	 * A EWS FindItem operation performed with a request constructed by this function will:
	 * - not search sub-folders: {@link ItemQueryTraversalType#SHALLOW}
	 * - return at most {@link #getMaxFindItems()}
	 * - returns {@link CalendarItemType}s but only the {@link ItemIdType} field is populated: {@link DefaultShapeNamesType#ID_ONLY}
	 * - searches using a {@link IndexedPageViewType} which means: 
	 * 		- results can be paged
 	 *		- results may contain recurring masters
 	 *
	 * This method
	 * @param restriction
	 * @param folderIds - if omitted the {@link FindItem} request will target the {@link DistinguishedFolderIdNameType#CALENDAR}
	 * @return
	 */
	protected FindItem constructIndexedPageViewFindFirstItemIdsShallow(RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		return constructIndexedPageViewFindItemIdsShallow(INIT_BASE_OFFSET, getMaxFindItems(), restriction, folderIds);
	}

	protected FindItem constructIndexedPageViewFindNextItemIdsShallow(int offset, RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		return constructIndexedPageViewFindItemIdsShallow(offset, getMaxFindItems(), restriction, folderIds);
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
		return constructIndexedPageViewFindItem(offset, maxItems, DefaultShapeNamesType.ID_ONLY, getItemQueryTraversal(), restriction, folderIds);
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
	private FindItem constructIndexedPageViewFindItem(int offset, int maxItems, DefaultShapeNamesType baseShape, ItemQueryTraversalType traversalType, RestrictionType restriction, Collection<? extends BaseFolderIdType> folderIds) {
		
		IndexedPageViewType view = constructIndexedPageView(offset,maxItems,false);
		
		ItemResponseShapeType responseShape = constructTextualItemResponseShapeInternal(baseShape);
		
		
		List<FieldOrderType> sortOrderList = null;
		FieldOrderType sortOrder = constructSortOrder();
		if(sortOrder != null ){
			sortOrderList = Collections.singletonList(sortOrder);
		}
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
			parentFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
			findItem.setParentFolderIds(parentFolderIds);
		}

		return findItem;
	}
	
	//================================================================================
    // DeleteItem
    //================================================================================	
	/**
	 * public DeleteItem constructDeleteItem
	 * @see https://msdn.microsoft.com/en-us/library/office/aa562961(v=exchg.140).aspx
	 * @see https://msdn.microsoft.com/en-us/library/office/bb204091(v=exchg.140).aspx
	 * @param itemIds
	 *            - Contains an array of items, occurrence items, and recurring
	 *            master items to delete from a mailbox in the Exchange store.
	 *            The DeleteItem Operation can be performed on any item type
	 * @param disposalType
	 *            - Describes how an item is deleted. This attribute is
	 *            required.
	 * @param sendMeetingInvites
	 *            - Describes whether a calendar item deletion is communicated
	 *            to attendees. This attribute is required when calendar items
	 *            are deleted. This attribute is optional if non-calendar items
	 *            are deleted.
	 * @param affectedTaskOccurrencesType
	 *            - Describes whether a task instance or a task master is
	 *            deleted by a DeleteItem Operation. This attribute is required
	 *            when tasks are deleted. This attribute is optional when
	 *            non-task items are deleted.
	 * @return {@link DeleteItem}
	 */
	private final  DeleteItem constructDeleteItemInternal(
			NonEmptyArrayOfBaseItemIdsType arrayOfItemIds,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendMeetingInvites,
			AffectedTaskOccurrencesType affectedTaskOccurrencesType) {
		DeleteItem deleteItem = new DeleteItem();
		deleteItem.setDeleteType(disposalType);
		deleteItem.setItemIds(arrayOfItemIds);
		if (null != affectedTaskOccurrencesType) {
			deleteItem.setAffectedTaskOccurrences(affectedTaskOccurrencesType);
		}
		if (null != sendMeetingInvites) {
			deleteItem.setSendMeetingCancellations(sendMeetingInvites);
		}
		return deleteItem;
	}
	
	/**
	 * Construct a {@link DeleteItem} request which will trigger the deletiion of the items specifiefied by {@code itemIds}
	 * @param itemIds
	 * @param disposalType
	 * @param sendMeetingInvites
	 * @param affectedTaskOccurrencesType
	 * @return
	 */
	protected DeleteItem constructDeleteItem(
			Collection<? extends BaseItemIdType> itemIds,
			DisposalType disposalType,
			CalendarItemCreateOrDeleteOperationType sendMeetingInvites,
			AffectedTaskOccurrencesType affectedTaskOccurrencesType) {
		Validate.notEmpty(itemIds, "must specify at least one itemId.");
		Validate.notNull(disposalType, "disposalType cannot be null");
		NonEmptyArrayOfBaseItemIdsType arrayOfItemIds = new NonEmptyArrayOfBaseItemIdsType();
		arrayOfItemIds
				.getItemIdsAndOccurrenceItemIdsAndRecurringMasterItemIds()
				.addAll(itemIds);
		return constructDeleteItemInternal(arrayOfItemIds, disposalType, sendMeetingInvites, affectedTaskOccurrencesType);
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
		if(null != exProps && !CollectionUtils.isEmpty(exProps.getPaths())){
			responseShape.setAdditionalProperties(exProps);
		}
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
	/**
	 * Construct a {@link FindFolder} request using an {@link IndexedPageViewType}, which can be used for paging.
	 * @param parent
	 * @param folderShape
	 * @param folderQueryTraversalType
	 * @param restriction
	 * @return
	 */
	protected final FindFolder constructFindFolderInternal(DistinguishedFolderIdNameType parent,
			DefaultShapeNamesType folderShape,
			FolderQueryTraversalType folderQueryTraversalType,
			RestrictionType restriction){
		
		FindFolder findFolder = new FindFolder();
		findFolder.setTraversal(folderQueryTraversalType);

		FolderResponseShapeType responseShape = new FolderResponseShapeType();
		responseShape.setBaseShape(folderShape);
		findFolder.setFolderShape(responseShape);

		IndexedPageViewType pageView = constructIndexedPageView(INIT_BASE_OFFSET, getMaxFindItems(), false);
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
	/**
	 * Construct an {@link UpdateFolder} request for deleting an extended property field.
	 * @param folderId the {@link FolderIdType} to update.
	 * @param exProp the {@link ExtendedPropertyType} to delete.
	 * @return {@link UpdateFolder}
	 */
	protected UpdateFolder constructUpdateFolderDeleteExtendedProperty(
			FolderIdType folderId, ExtendedPropertyType exProp) {
		return constructUpdateFolderDeleteField(
				getJaxbPathToExtendedFieldType(exProp), folderId);
	}

	/**
	 * Construct an {@link UpdateFolder} request for setting a property field.
	 * @param folder the {@link FolderType} contains the changes.
	 * @param path the {@link JAXBElement} idenifies the path to the property field that will be updated.
	 * @param folderId the {@link FolderIdType} to update.
	 * @return {@link UpdateFolder}
	 */
	protected UpdateFolder constructUpdateFolderSetField(FolderType folder,
			JAXBElement<? extends BasePathToElementType> path,
			FolderIdType folderId) {
		SetFolderFieldType changeDescription = new SetFolderFieldType();
		changeDescription.setFolder(folder);
		changeDescription.setPath(path);
		return constructUpdateFolderInternal(changeDescription, folderId);
	}
	/**
	 * Construct an {@link UpdateFolder} request for deleting a property field.
	 * @param path the {@link JAXBElement} idenifies the path to the property field that will be deleted.
	 * @param folderId the {@link FolderIdType} to update.
	 * @return {@link UpdateFolder}
	 */
	protected UpdateFolder constructUpdateFolderDeleteField(
			JAXBElement<? extends BasePathToElementType> path,
			FolderIdType folderId) {
		DeleteFolderFieldType changeDescription = new DeleteFolderFieldType();
		changeDescription.setPath(path);
		return constructUpdateFolderInternal(changeDescription, folderId);
	}

	/**
	 * @param changeDescription
	 * @param folderId
	 * @return
	 */
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
	protected final EmptyFolder constructEmptyFolder(boolean deleteSubFolders, DisposalType disposalType, Collection<? extends BaseFolderIdType> folderIds){
		EmptyFolder request = new EmptyFolder();
		request.setDeleteSubFolders(deleteSubFolders);
		request.setDeleteType(disposalType);
		NonEmptyArrayOfBaseFolderIdsType nonEmptyArrayOfBaseFolderIds = new NonEmptyArrayOfBaseFolderIdsType();
		nonEmptyArrayOfBaseFolderIds.getFolderIdsAndDistinguishedFolderIds().addAll(folderIds);
		request.setFolderIds(nonEmptyArrayOfBaseFolderIds);
		return request;
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
	private final CalendarViewType constructCalendarView(Date startTime, Date endTime) {
		CalendarViewType calendarView = new CalendarViewType();
		calendarView.setMaxEntriesReturned(getMaxFindItems());
		calendarView.setStartDate(ExchangeDateUtils.convertDateToXMLGregorianCalendar(startTime));
		calendarView.setEndDate(ExchangeDateUtils.convertDateToXMLGregorianCalendar(endTime));
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
	private final RestrictionType constructAndRestriction(Collection<JAXBElement<? extends SearchExpressionType>> expressions){
		AndType andType = new AndType();
		andType.getSearchExpressions().addAll(expressions);
		JAXBElement<AndType> andSearchExpression = getObjectFactory().createAnd(andType);
		
		// create restriction and set searchExpression
		RestrictionType restrictionType = new RestrictionType();
		restrictionType.setSearchExpression(andSearchExpression);
		return restrictionType;
	}
	/**
	 * Construct a {@link RestrictionType} object for use with a {@link FindItem} using a {@link IndexedPageViewType}.  
	 * {@link RestrictionType} cannot be used with {@link CalendarViewType}.
	 * 
	 * https://tools.ietf.org/html/rfc4791#section-9.9
	 * 
	 * @param startTime
	 * @param endTime
	 * @return a {@link RestrictionType} which targets {@link CalendarItemType} betweem {@code startTime} and {@code endTime}
	 */
	protected RestrictionType constructFindCalendarItemsByDateRangeRestriction(Date startTime, Date endTime) {
		Collection<JAXBElement<? extends SearchExpressionType>> expressions = new HashSet<JAXBElement<? extends SearchExpressionType>>();
		expressions.add(getCalendarItemStartSearchExpression(startTime));
		expressions.add(getCalendarItemEndSearchExpression(endTime));
		return constructAndRestriction(expressions);
	}
	
	protected RestrictionType constructFindCancelledCalendarItemsByDateRangeRestriction(Date startTime, Date endTime){
		Collection<JAXBElement<? extends SearchExpressionType>> expressions = new HashSet<JAXBElement<? extends SearchExpressionType>>();
		expressions.add(getCalendarItemStartSearchExpression(startTime));
		expressions.add(getCalendarItemEndSearchExpression(endTime));
		expressions.add(getCalendarItemIsCancelledSearchExpression());
		return constructAndRestriction(expressions);
	}
	
	
	protected RestrictionType constructFindCancelledCalendarItemsRestriction(){
		RestrictionType restriction = new RestrictionType();
		restriction.setSearchExpression(getCalendarItemIsCancelledSearchExpression());
		return restriction;
	}
	
	/**
	 * @see https://msdn.microsoft.com/en-us/library/ee202361(v=exchg.80).aspx
	 * @return
	 */
	protected JAXBElement<IsGreaterThanOrEqualToType> getCalendarItemIsCancelledSearchExpression(){
		ConstantValueType isCancelledValue = new ConstantValueType();
		isCancelledValue.setValue("4");
		
		FieldURIOrConstantType constant = new FieldURIOrConstantType();
		constant.setConstant(isCancelledValue);
		
		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
		//path.setFieldURI(UnindexedFieldURIType.CALENDAR_IS_CANCELLED);
		path.setFieldURI(UnindexedFieldURIType.CALENDAR_APPOINTMENT_STATE);
		
		IsGreaterThanOrEqualToType equalToType = new IsGreaterThanOrEqualToType();	
		equalToType.setFieldURIOrConstant(constant);
		equalToType.setPath(getObjectFactory().createFieldURI(path));
		
		return getObjectFactory().createIsGreaterThanOrEqualTo(equalToType);
	}
	
	/**
	 * Construct a search expression for finding calendar itmems that start BEFORE the specified {@code endTime}
	 * @param endTime
	 * @return {@link JAXBElement<IsLessThanOrEqualToType>}
	 */
	protected final JAXBElement<IsLessThanType> getCalendarItemEndSearchExpression(Date endTime) {
		ObjectFactory of = getObjectFactory();
		IsLessThanType endType = new IsLessThanType();
		XMLGregorianCalendar end = ExchangeDateUtils.convertDateToXMLGregorianCalendar(endTime);
		PathToUnindexedFieldType endPath = new PathToUnindexedFieldType();
		//endPath.setFieldURI(UnindexedFieldURIType.CALENDAR_END);
		endPath.setFieldURI(UnindexedFieldURIType.CALENDAR_START);
		JAXBElement<PathToUnindexedFieldType> endFieldURI = of.createFieldURI(endPath);
		endType.setPath(endFieldURI);
		FieldURIOrConstantType endConstant = new FieldURIOrConstantType();
		ConstantValueType endValue = new ConstantValueType();
		endValue.setValue(end.toXMLFormat());
		endConstant.setConstant(endValue);
		endType.setFieldURIOrConstant(endConstant);
		JAXBElement<IsLessThanType> endSearchExpression = of.createIsLessThan(endType);
		return endSearchExpression;
	}
	
	/**
	 * Construct a search expression for finding calendar itmems ON or AFTER the specified {@code startTime}
	 * @param endTime
	 * @return {@link JAXBElement<IsGreaterThanOrEqualToType>}
	 */
	protected final JAXBElement<IsGreaterThanOrEqualToType> getCalendarItemStartSearchExpression(Date startTime) {
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
		//TODO this is wrong ItemResponseShape is only used with FindItem.  Should a findItem request return the extendedProperties identified by getItemExtendedPropertyPaths()?
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
	 * Warning!  applying a sortOrder to a FindItem operation can have inconsistent results (errorInternalServerError when targetting aims, NO_ERROR when targeting O365....)
	 * 
	 * @return
	 */
	protected FieldOrderType constructSortOrder() {
		return null;
//		ObjectFactory of = getObjectFactory();
//		//set sort order (earliest items first)
//		FieldOrderType sortOrder = new FieldOrderType();
//		sortOrder.setOrder(SortDirectionType.ASCENDING);
//		PathToUnindexedFieldType path = new PathToUnindexedFieldType();
//		path.setFieldURI(UnindexedFieldURIType.ITEM_ITEM_ID);
//		//path.setFieldURI(UnindexedFieldURIType.CALENDAR_START);
//		JAXBElement<PathToUnindexedFieldType> sortPath = of.createFieldURI(path);
//		sortOrder.setPath(sortPath);
//		return sortOrder;
	}
}
