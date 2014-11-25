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

import java.util.List;
import java.util.Set;

import org.apache.commons.lang3.tuple.Pair;

import com.microsoft.exchange.messages.BaseResponseMessageType;
import com.microsoft.exchange.messages.CreateFolderResponse;
import com.microsoft.exchange.messages.CreateItemResponse;
import com.microsoft.exchange.messages.DeleteFolderResponse;
import com.microsoft.exchange.messages.EmptyFolderResponse;
import com.microsoft.exchange.messages.FindFolderResponse;
import com.microsoft.exchange.messages.FindItemResponse;
import com.microsoft.exchange.messages.GetFolderResponse;
import com.microsoft.exchange.messages.GetItemResponse;
import com.microsoft.exchange.messages.GetServerTimeZonesResponse;
import com.microsoft.exchange.messages.ResolveNamesResponse;
import com.microsoft.exchange.messages.ResponseCodeType;
import com.microsoft.exchange.messages.UpdateFolderResponse;
import com.microsoft.exchange.messages.UpdateItemResponse;
import com.microsoft.exchange.types.BaseFolderType;
import com.microsoft.exchange.types.CalendarItemType;
import com.microsoft.exchange.types.FolderIdType;
import com.microsoft.exchange.types.ItemIdType;
import com.microsoft.exchange.types.ItemType;
import com.microsoft.exchange.types.ResponseClassType;
import com.microsoft.exchange.types.TimeZoneDefinitionType;

/**
 * @author ctcudd
 *
 */
public interface ExchangeResponseUtils {

	//================================================================================
    // Folder Responses
    //================================================================================	
	/**
	 * Parse a {@link CreateFolderResponse} message and return a {@link Set} of {@link FolderIdType}s representing the folders that were created.
	 * @param response
	 * @return - a never null but possibly empty {@link Set} of {@link FolderIdType}s
	 */
	public Set<FolderIdType> parseCreateFolderResponse(CreateFolderResponse response);

	/**
	 * Parse a {@link UpdateFolderResponse} message and return a {@link Set} of {@link FolderIdType}s representing the folders that were updated.
	 * @param response
	 * @return - a never null but possibly empty {@link Set} of {@link FolderIdType}s
	 */
	public Set<FolderIdType> parseUpdateFolderResponse(UpdateFolderResponse response);

	/**
	 * Parse a {@link GetFolderResponse} message and return a {@link Set} of {@link BaseFolderType}s
	 * @param getFolderResponse
	 * @return a never null but possibly empty {@link Set} of {@link BaseFolderType}s
	 */
	public Set<BaseFolderType> parseGetFolderResponse(GetFolderResponse getFolderResponse);
	
	/**
	 * Parse a {@link FindFolderResponse} message and return a {@link Set} of {@link BaseFolderType}s
	 * @param getFolderResponse
	 * @return a never null but possibly empty {@link Set} of {@link BaseFolderType}s
	 */
	public Set<BaseFolderType> parseFindFolderResponse(FindFolderResponse findFolderResponse);
	
	/**
	 * Parse an {@link EmptyFolderResponse} and return a boolean indicating whether the operation was successful or not
	 * @param response
	 * @return whether the operation was successful or not
	 */
	public boolean parseEmptyFolderResponse(EmptyFolderResponse response);

	/**
	 * Parse an {@link DeleteFolderResponse} and return a boolean indicating whether the operation was successful or not
	 * @param response
	 * @return whether the operation was successful or not
	 */
	public boolean parseDeleteFolderResponse(DeleteFolderResponse response);
	
	//================================================================================
    // Item Responses
    //================================================================================	

	/**
	 * Parse a {@link CreateItemResponse} and return only the {@link ItemIdType}s that correspond to successfully created items
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}s
	 */
	public Set<ItemIdType> parseCreateItemResponse(CreateItemResponse response);
	
	/**
	 * Parse a {@link CreateItemResponse} and return only the {@link ItemIdType}s that correspond to successfully created items
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}s
	 */
	public Set<ItemIdType> getCreatedItemIds(CreateItemResponse response);

	/**
	 * Not sure what this is for
	 * @param response
	 * @return
	 */
	@Deprecated
	public List<String> getCreateItemErrors(CreateItemResponse response);

	/**
	 * Parse a {@link GetItemResponse} message and return a {@link Set} of {@link ItemType}s.
	 * @param response
	 * @return - a never null but possibly empty {@link Set} of {@link ItemType}s
	 */
	public Set<ItemType> parseGetItemResponse(GetItemResponse response);

	/**
	 * Parse a {@link FindItemResponse} and return a {@link Set} of {@link ItemIdType} corresponding to the found {@link ItemType}s
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}s
	 */
	public Set<ItemType> parseFindItemResponse(FindItemResponse response);

	/**
	 * Parse a {@link FindItemResponse} and return a {@link Set} of {@link CalendarItemType}s
	 * This method will omit any non calendar items from the result set
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link CalendarItemType}s
	 */
	public Set<CalendarItemType> parseFindCalendarItemResponse( FindItemResponse response);
	
	/**
	 * Parse a {@link FindItemResponse} and return a {@link Set} of {@link ItemIdType} corresponding to the found {@link ItemType}s
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}s
	 */
	public Set<ItemIdType> parseFindItemIdResponseNoOffset(FindItemResponse response);

	/**
	 * Parse a {@link FindItemResponse} and return a {@link Pair} where:
	 *  the left side of the result is a {@link Set} of {@link ItemIdType}s corresponding to the found {@link ItemType}s
	 *  and the right side is an integer representing the offset of the lastItem returned.
	 *  
	 *  This method is especially useful for parsing paged results.
	 *  
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link ItemIdType}s
	 */
	public Pair<Set<ItemIdType>, Integer> parseFindItemIdResponse(FindItemResponse response);
	
	/**
	 * Parse an {@link UpdateItemResponse} message and return a {@link Set} of {@link ItemIdType} corresponding to the {@link ItemType}s that were updated
	 * @param response
	 * @return
	 */
	public Set<ItemIdType> parseUpdateItemResponse(UpdateItemResponse response);

	
	//================================================================================
    // Misc Responses
    //================================================================================
	/**
	 * Parse a {@link ResolveNamesResponse} and
	 * @param response
	 * @return a never null but possibly empty {@link Set} of {@link String}s representing the email address associated with the resolveNames request
	 */
	public Set<String> parseResolveNamesResponse(ResolveNamesResponse response);

	/**
	 * Parse a {@link GetServerTimeZonesResponse} and return a {@link Set} of {@link TimeZoneDefinitionType}s
	 * @param response
	 * @return a never null but possible empty {@link Set} of {@link TimeZoneDefinitionType}
	 */
	public Set<TimeZoneDefinitionType> parseGetServerTimeZonesResponse( GetServerTimeZonesResponse response);
	
	//================================================================================
    // BaseResponseMessageType Responses
    //================================================================================
	/**
	 * @param response
	 * @return true if the response has {@link ResponseCodeType#NO_ERROR}, false
	 *         otherwise
	 */
	public boolean confirmSuccess(BaseResponseMessageType response);
	
	/**
	 * @param response
	 * @return true if the response contains {@link ResponseClassType#SUCCESS} or {@link ResponseClassType#WARNING}
	 */
	public boolean confirmSuccessOrWarning(BaseResponseMessageType response);

}