/**
 * 
 */
package com.microsoft.exchange.impl;

import java.util.Collection;
import java.util.Date;
import java.util.Set;

import org.springframework.retry.annotation.Backoff;
import org.springframework.retry.annotation.Retryable;
import org.springframework.stereotype.Service;

import com.microsoft.exchange.ExchangeRequestFactory;
import com.microsoft.exchange.exception.ExchangeWebServicesRuntimeException;
import com.microsoft.exchange.messages.FindItem;
import com.microsoft.exchange.types.BaseFolderIdType;
import com.microsoft.exchange.types.CalendarFolderType;
import com.microsoft.exchange.types.CalendarItemCreateOrDeleteOperationType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.DisposalType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.ItemIdType;

/**
 * @author Collin Cudd
 *
 */
@Service
public interface ExchangeCalendarDataDao {

	/**
	 * Find all {@link ItemIdType} within the primary {@link CalendarFolderType} between {@code startDate} and {@code endDate}
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	public abstract Set<ItemIdType> findCalendarItemIds(String upn,
			Date startDate, Date endDate);

	/**
	 * Find all {@link ItemIdType} within the specified {@link FolderIdType}s between {@code startDate} and {@code endDate}
	 * @param upn
	 * @param startDate
	 * @param endDate
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	public abstract Set<ItemIdType> findCalendarItemIds(String upn,
			Date startDate, Date endDate, Collection<FolderIdType> calendarIds);

	/**
	 * Obtain all {@link ItemIdType}s within a specified {@link FolderIdType} by
	 * repeatedly callling {@link FindItem} and paging the results
	 * 
	 * @param upn
	 * @param folderIds
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}
	 */
	public abstract Set<ItemIdType> findAllItemIds(String upn,
			Collection<FolderIdType> folderIds);

	/**
	 * Create the {@link CalendarItemType} on the exchange server
	 * @param upn
	 * @param calendarItem
	 * @return {@link ItemIdType}
	 */
	public abstract ItemIdType createCalendarItem(String upn,
			CalendarItemType calendarItem);

	/**
	 * Create a {@link CalendarFolderType} with the specified {@code displayName}
	 * @param upn
	 * @param displayName
	 * @return {@link FolderIdType}
	 */
	public abstract FolderIdType createCalendarFolder(String upn,
			String displayName);

	/**
	 * Delete {@link CalendarItemType}s for the user.
	 * @param upn - the userPrincipalName identifies the user to delete {@link CalendarItemType}s for.
	 * @param itemIds - the {@link ItemIdType}s for the {@link CalendarItemType}s to delete.
	 * 
	 * @see ExchangeRequestFactory#constructDeleteCalendarItems(Collection, DisposalType, CalendarItemCreateOrDeleteOperationType)
	 * @return
	 */
	@Retryable(maxAttempts=10, include=ExchangeWebServicesRuntimeException.class, backoff=@Backoff(delay=100, multiplier=2, random=true))
	public abstract boolean deleteCalendarItemsWithRetry(String upn,
			Collection<ItemIdType> itemIds);

	public abstract boolean deleteCalendarItems(String upn,
			Collection<ItemIdType> itemIds);

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
	public abstract boolean deleteCalendarFolder(String upn,
			FolderIdType folderId);

	/**
	 * Delete a calendar folder
	 * @param upn - the user 
	 * @param disposalType - how the deletion is performed
	 * @param folderId - the folder to delete
	 * @return
	 */
	public abstract boolean deleteFolder(String upn, BaseFolderIdType folderId);

}