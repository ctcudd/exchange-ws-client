/**
 * 
 */
package com.microsoft.exchange.impl;
import org.apache.commons.lang.StringUtils;
import org.springframework.dao.support.DataAccessUtils;

import com.microsoft.exchange.ExchangeRequestFactory;
import com.microsoft.exchange.messages.CreateItem;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.BaseItemIdType;
import com.microsoft.exchange.types.BodyTypeResponseType;
import com.microsoft.exchange.types.CalendarViewType;
import com.microsoft.exchange.types.DefaultShapeNamesType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.IndexBasePointType;
import com.microsoft.exchange.types.IndexedPageViewType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.ItemQueryTraversalType;
import com.microsoft.exchange.types.ItemResponseShapeType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.NonEmptyArrayOfAllItemsType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseFolderIdsType;
import com.microsoft.exchange.types.NonEmptyArrayOfBaseItemIdsType;
import com.microsoft.exchange.types.WellKnownResponseObjectType;

/**
 * Utility methods used only for testing {@link ExchangeRequestFactory}.
 * 
 * This class is intended to aid in testing {@link ExchangeRequestFactory}
 * methods. This class is in main, the intent is that any project extending
 * {@link ExchangeRequestFactory} can and should write a unit test class which
 * extends {@link ExchangeRequestFactoryUtils} should help.
 * 
 * TODO consider moving this to a seperate module: exchange-ws-client-test ?
 * 
 * @author Collin Cudd
 *
 */
public abstract class ExchangeRequestFactoryUtils {

	public abstract ExchangeRequestFactory getFactory();
	
	protected FolderIdType createFolderId(String folderId){
		return createFolderId(folderId, null);
	}
	
	protected FolderIdType createFolderId(){
		return createFolderId(null, null);
	}
	/**
	 * @return
	 */
	protected FolderIdType createFolderId(String folderId, String changeKey) {
		FolderIdType f = new FolderIdType();
		if(StringUtils.isBlank(folderId)){
			f.setId("itemid");
		}else{
			f.setId(folderId);
		}
		return f;
	}
	
	protected ItemIdType createItemId() {
		return createItemId(null,null);
	}
	
	protected ItemIdType createItemId(String itemId) {
		return createItemId(itemId,null);
	}
	
	protected ItemIdType createItemId(String itemId, String changeKey) {
		ItemIdType i = new ItemIdType();
		if(StringUtils.isBlank(itemId)){
			i.setId("itemid");
		}else{
			i.setId(itemId);
		}
		if(StringUtils.isBlank(changeKey)){
			i.setChangeKey("changeKey");
		}else{
			i.setChangeKey(changeKey);
		}
		return i;
	}
	
	/**
	 * @param request a {@link FindItem} request
	 * @return a {@link CalendarViewType} if the {@link FindItem} request contains no other view, null otherwise.
	 */
	protected CalendarViewType findItemIsCalendarView(FindItem request) {
		if(null == request){
			return null;
		}
		if(null != request.getIndexedPageItemView()) {
			return null;
		}
		if(null != request.getContactsView()){
			return null;
		}
		if(null != request.getFractionalPageItemView()){
			return null;
		}
		return request.getCalendarView();
	}
	
	/**
	 * @param request a {@link FindItem} request
	 * @return a {@link IndexedPageViewType} if the {@link FindItem} request contains no other view, null otherwise.
	 */
	protected IndexedPageViewType findItemIsIndexedPageView(FindItem request){
		if(null == request){
			return null;
		}
		if(null != request.getCalendarView()) {
			return null;
		}
		if(null != request.getContactsView()){
			return null;
		}
		if(null != request.getFractionalPageItemView()){
			return null;
		}
		return request.getIndexedPageItemView();
	}
	
	/**
	 * @param shape
	 * @param expectedBaseShape
	 * @param expectedBodyType
	 * @return true if and only if the shape provided contains a matching {@link DefaultShapeNamesType} and {@link BodyTypeResponseType}
	 */
	protected boolean checkResponseShape(ItemResponseShapeType shape, DefaultShapeNamesType expectedBaseShape, BodyTypeResponseType expectedBodyType){
		if(null == shape) return false;
		if(!expectedBaseShape.equals(shape.getBaseShape())){
			return false;
		}
		if(!expectedBodyType.equals(shape.getBodyType())){
			return false;
		}
		return true;
	}
	
	/**
	 * 
	 * @param request
	 * @return a {@link WellKnownResponseObjectType} only if the
	 *         {@link CreateItem} request contains a single item which is an
	 *         instance of {@link WellKnownResponseObjectType}.
	 */
	public WellKnownResponseObjectType createItemIsWellKnownResponseType(CreateItem request){
		if(request == null){
			return null;
		}
		ItemType item = createItemContainsSingleItem(request);
		if(null == item){
			return null;
		}
		if(item instanceof WellKnownResponseObjectType){
			return ((WellKnownResponseObjectType) item);
		}
		return null;
	}
	
	/**
	 * @param request the {@link CreateItem} request
	 * @return {@link ItemType} or null.
	 */
	public ItemType createItemContainsSingleItem(CreateItem request){
		if(request == null){
			return null;
		}
		return containsSingleItem(request.getItems());
	}
	
	public ItemType containsSingleItem(NonEmptyArrayOfAllItemsType itemsArray){
		if(null == itemsArray){
			return null;
		}
		return DataAccessUtils.singleResult(itemsArray.getItemsAndMessagesAndCalendarItems());
	}
	
	public BaseItemIdType containsSingleItemId(NonEmptyArrayOfBaseItemIdsType itemIds){
		if(null == itemIds){
			return null;
		}
		return DataAccessUtils.singleResult(itemIds.getItemIdsAndOccurrenceItemIdsAndRecurringMasterItemIds());
	}
	
	public BaseFolderIdType containsSingleFolderId(NonEmptyArrayOfBaseFolderIdsType folderIds){
		if(null == folderIds){
			return null;
		}
		return DataAccessUtils.singleResult(folderIds.getFolderIdsAndDistinguishedFolderIds());
	}
	/**
	 * @param traversal
	 * @param expectedTraversalType
	 * @return true if and only if the traversal provided contains a matching {@link ItemQueryTraversalType}.
	 */
	protected boolean checkTraversal(ItemQueryTraversalType traversal, ItemQueryTraversalType expectedTraversalType){
		if(null == traversal){
			return false;
		}
		if(!expectedTraversalType.equals(traversal)){
			return false;
		}
		return true;
	}
	
	protected boolean indexedPageViewTypeBasePointIsBeginning(IndexedPageViewType view){
		//the IndexedPageViewType should start at the beginning, i.e. 0, and return at most factory.getMaxFindItems()
		if(view == null){
			return false;
		}
		if(!IndexBasePointType.BEGINNING.equals(view.getBasePoint())){
			return false;
		}
		if(0 != view.getOffset()){
			return false;
		}
		if(getFactory().getMaxFindItems() != view.getMaxEntriesReturned().intValue()){
			return false;
		}
		return true;
	}
}
