"""
Outlook Model - Data layer for Outlook COM interactions
Enhanced with comprehensive error handling and logging
"""

import win32com.client
from collections import defaultdict
from datetime import datetime
from typing import List, Dict, Optional, Tuple
import logging
import traceback


# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('outlook_reader.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class OutlookConnectionError(Exception):
    """Raised when connection to Outlook fails"""
    pass


class OutlookDataError(Exception):
    """Raised when reading data from Outlook fails"""
    pass


class OutlookModel:
    """Model class for interacting with Microsoft Outlook via COM"""

    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.inbox = None
        self._connected = False
        logger.info("OutlookModel initialized")

    def connect(self) -> Tuple[bool, str]:
        """
        Connect to Microsoft Outlook application.
        Returns: (success: bool, message: str)
        """
        logger.info("Attempting to connect to Outlook...")

        try:
            # Attempt to connect to Outlook
            logger.debug("Dispatching Outlook.Application COM object...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")

            if not self.outlook:
                raise OutlookConnectionError("Failed to create Outlook COM object")

            logger.debug("Getting MAPI namespace...")
            self.namespace = self.outlook.GetNamespace("MAPI")

            if not self.namespace:
                raise OutlookConnectionError("Failed to get MAPI namespace")

            logger.debug("Accessing Inbox folder (folder index 6)...")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox

            if not self.inbox:
                raise OutlookConnectionError("Failed to access Inbox folder")

            # Test that we can actually read from inbox
            try:
                item_count = self.inbox.Items.Count
                logger.info(f"Successfully connected to Outlook. Inbox contains {item_count} items.")
            except Exception as e:
                logger.error(f"Connected but cannot read inbox items: {e}")
                raise OutlookConnectionError(f"Cannot access inbox items: {str(e)}")

            self._connected = True
            return True, f"Successfully connected to Outlook ({item_count} emails in inbox)"

        except Exception as e:
            self._connected = False
            error_msg = f"Failed to connect to Outlook: {str(e)}"
            logger.error(error_msg)
            logger.debug(f"Connection error traceback: {traceback.format_exc()}")

            # Provide specific error messages for common issues
            if "invalid class string" in str(e).lower():
                error_msg = "Outlook is not installed or not properly registered"
            elif "access denied" in str(e).lower():
                error_msg = "Access denied. Please ensure Outlook is configured and you have permissions"
            elif "rpc server" in str(e).lower():
                error_msg = "Cannot communicate with Outlook. Please ensure Outlook is running"

            return False, error_msg

    def is_connected(self) -> bool:
        """Check if connected to Outlook"""
        return self._connected

    def get_inbox_count(self) -> int:
        """Get total number of messages in inbox"""
        if not self._connected or not self.inbox:
            logger.warning("Cannot get inbox count - not connected")
            return 0

        try:
            count = self.inbox.Items.Count
            logger.debug(f"Inbox contains {count} items")
            return count
        except Exception as e:
            logger.error(f"Error getting inbox count: {e}")
            return 0

    def get_conversations(self) -> List[Dict]:
        """
        Get all inbox messages grouped by conversation.
        Returns list of conversation dictionaries.
        """
        logger.info("Starting to retrieve conversations...")

        if not self._connected:
            error_msg = "Cannot get conversations - not connected to Outlook"
            logger.error(error_msg)
            raise OutlookDataError(error_msg)

        if not self.inbox:
            error_msg = "Cannot get conversations - inbox not initialized"
            logger.error(error_msg)
            raise OutlookDataError(error_msg)

        try:
            # Get all messages
            logger.debug("Accessing inbox items...")
            messages = self.inbox.Items

            if not messages:
                logger.warning("Inbox.Items is None or empty")
                return []

            message_count = messages.Count
            logger.info(f"Found {message_count} messages in inbox")

            if message_count == 0:
                logger.info("Inbox is empty - returning empty conversation list")
                return []

            conversations = defaultdict(list)
            processed_count = 0
            error_count = 0

            # Process each message
            logger.debug(f"Processing {message_count} messages...")
            for i, message in enumerate(messages, 1):
                try:
                    # Log progress every 50 messages
                    if i % 50 == 0:
                        logger.debug(f"Processing message {i}/{message_count}")

                    # Check message class (only process mail items)
                    try:
                        msg_class = message.Class
                        # 43 = olMail (standard email message)
                        if msg_class != 43:
                            logger.debug(f"Skipping non-email item (class={msg_class})")
                            continue
                    except Exception as e:
                        logger.warning(f"Cannot get message class, skipping: {e}")
                        continue

                    # Extract message properties safely
                    try:
                        subject = self._safe_get_property(message, 'Subject', "(No Subject)")
                    except Exception as e:
                        logger.warning(f"Error getting subject: {e}")
                        subject = "(No Subject)"

                    try:
                        conv_id = self._safe_get_property(message, 'ConversationID', None)
                    except Exception as e:
                        logger.debug(f"No ConversationID available: {e}")
                        conv_id = None

                    try:
                        received_time = self._safe_get_property(message, 'ReceivedTime', None)
                    except Exception as e:
                        logger.warning(f"Error getting received time: {e}")
                        received_time = None

                    try:
                        sender = self._safe_get_property(message, 'SenderName', "Unknown")
                    except Exception as e:
                        logger.warning(f"Error getting sender: {e}")
                        sender = "Unknown"

                    try:
                        sender_email = self._safe_get_property(message, 'SenderEmailAddress', "")
                    except Exception as e:
                        logger.debug(f"Error getting sender email: {e}")
                        sender_email = ""

                    try:
                        unread = self._safe_get_property(message, 'UnRead', False)
                    except Exception as e:
                        logger.debug(f"Error getting unread status: {e}")
                        unread = False

                    # Use conversation ID or unique ID as grouping key
                    if conv_id:
                        group_key = conv_id
                    else:
                        # Create unique key for non-threaded emails
                        try:
                            entry_id = self._safe_get_property(message, 'EntryID', None)
                            group_key = f"single_{entry_id}" if entry_id else f"single_{i}"
                        except:
                            group_key = f"single_{i}"

                    # Add to conversations
                    email_data = {
                        'subject': subject,
                        'sender': sender,
                        'sender_email': sender_email,
                        'received_time': received_time,
                        'conv_id': conv_id,
                        'unread': unread
                    }

                    conversations[group_key].append(email_data)
                    processed_count += 1

                    logger.debug(f"Processed email: '{subject[:50]}...' from {sender}")

                except Exception as e:
                    error_count += 1
                    logger.error(f"Error processing message {i}: {e}")
                    logger.debug(f"Message processing error traceback: {traceback.format_exc()}")
                    continue

            logger.info(f"Processed {processed_count}/{message_count} messages successfully ({error_count} errors)")
            logger.info(f"Grouped into {len(conversations)} conversation(s)")

            if processed_count == 0:
                logger.warning("No messages were successfully processed")
                return []

            # Sort emails within each conversation by time (oldest first)
            logger.debug("Sorting emails within conversations...")
            for conv_id in conversations:
                try:
                    conversations[conv_id].sort(
                        key=lambda x: x['received_time'] if x['received_time'] else datetime.min
                    )
                except Exception as e:
                    logger.error(f"Error sorting conversation {conv_id}: {e}")

            # Convert to list and sort conversations by most recent message
            logger.debug("Building conversation list...")
            conversation_list = []

            for conv_id, emails in conversations.items():
                try:
                    # Get latest time
                    valid_times = [email['received_time'] for email in emails if email['received_time']]
                    latest_time = max(valid_times) if valid_times else datetime.min

                    # Check for unread
                    has_unread = any(email.get('unread', False) for email in emails)

                    # Get subject
                    subject = emails[0]['subject'] if emails else "(No Subject)"

                    conversation_list.append({
                        'conv_id': conv_id,
                        'emails': emails,
                        'latest_time': latest_time,
                        'count': len(emails),
                        'has_unread': has_unread,
                        'subject': subject
                    })

                except Exception as e:
                    logger.error(f"Error building conversation entry: {e}")
                    continue

            # Sort by latest message time (newest first)
            logger.debug("Sorting conversations by latest time...")
            try:
                conversation_list.sort(key=lambda x: x['latest_time'], reverse=True)
            except Exception as e:
                logger.error(f"Error sorting conversations: {e}")

            logger.info(f"Successfully built {len(conversation_list)} conversation(s)")

            # Log summary
            total_emails = sum(conv['count'] for conv in conversation_list)
            unread_convs = sum(1 for conv in conversation_list if conv['has_unread'])
            logger.info(f"Summary: {total_emails} emails, {len(conversation_list)} conversations, {unread_convs} with unread")

            return conversation_list

        except OutlookDataError:
            # Re-raise our custom exceptions
            raise
        except Exception as e:
            error_msg = f"Unexpected error reading conversations: {str(e)}"
            logger.error(error_msg)
            logger.debug(f"Error traceback: {traceback.format_exc()}")
            raise OutlookDataError(error_msg)

    def _safe_get_property(self, obj, prop_name: str, default=None):
        """
        Safely get a property from a COM object.
        Returns default if property doesn't exist or fails.
        """
        try:
            if hasattr(obj, prop_name):
                value = getattr(obj, prop_name)
                # Handle None/empty values
                if value is None:
                    return default
                return value
            else:
                return default
        except Exception as e:
            logger.debug(f"Cannot get property '{prop_name}': {e}")
            return default

    def search_conversations(self, query: str) -> List[Dict]:
        """
        Search conversations by subject or sender.
        Returns filtered list of conversations.
        """
        logger.info(f"Searching conversations for: '{query}'")

        try:
            if not query or not query.strip():
                logger.debug("Empty query - returning all conversations")
                return self.get_conversations()

            all_conversations = self.get_conversations()
            query_lower = query.lower().strip()

            filtered = []
            for conv in all_conversations:
                try:
                    # Search in subject
                    if query_lower in conv['subject'].lower():
                        filtered.append(conv)
                        continue

                    # Search in sender names
                    for email in conv['emails']:
                        if query_lower in email['sender'].lower():
                            filtered.append(conv)
                            break
                except Exception as e:
                    logger.error(f"Error filtering conversation: {e}")
                    continue

            logger.info(f"Found {len(filtered)} matching conversation(s)")
            return filtered

        except Exception as e:
            logger.error(f"Error in search: {e}")
            raise OutlookDataError(f"Search failed: {str(e)}")

    def get_folder_list(self) -> List[str]:
        """Get list of available Outlook folders"""
        logger.debug("Getting folder list...")

        if not self._connected or not self.namespace:
            logger.warning("Cannot get folders - not connected")
            return []

        try:
            folders = []
            for folder in self.namespace.Folders:
                try:
                    folder_name = folder.Name
                    folders.append(folder_name)
                    logger.debug(f"Found folder: {folder_name}")
                except Exception as e:
                    logger.warning(f"Error reading folder: {e}")
                    continue

            logger.info(f"Found {len(folders)} folder(s)")
            return folders

        except Exception as e:
            logger.error(f"Error getting folder list: {e}")
            return []

    def disconnect(self):
        """Disconnect from Outlook"""
        logger.info("Disconnecting from Outlook...")
        try:
            self.outlook = None
            self.namespace = None
            self.inbox = None
            self._connected = False
            logger.info("Disconnected successfully")
        except Exception as e:
            logger.error(f"Error during disconnect: {e}")
