"""
Outlook Inbox Reader - Main Application
Modern GUI application with proper MVC architecture
Enhanced with comprehensive error handling
"""

import threading
import logging
import traceback
from typing import List, Dict
from models.outlook_model import OutlookModel, OutlookConnectionError, OutlookDataError
from views.main_window import MainWindow

logger = logging.getLogger(__name__)


class OutlookInboxApp:
    """Main application controller - coordinates Model and View"""

    def __init__(self):
        logger.info("="*60)
        logger.info("Outlook Inbox Reader - Application Starting")
        logger.info("="*60)

        try:
            # Initialize model and view
            logger.info("Initializing model and view...")
            self.model = OutlookModel()
            self.view = MainWindow()

            # Store current conversations
            self.current_conversations: List[Dict] = []

            # Setup view callbacks
            logger.info("Setting up view callbacks...")
            self.view.on_refresh_callback = self.refresh_conversations
            self.view.on_search_callback = self.search_conversations

            # Auto-connect on startup
            logger.info("Scheduling auto-connect...")
            self.view.after(500, self.auto_connect)

            logger.info("Application initialized successfully")

        except Exception as e:
            logger.error(f"Critical error initializing application: {e}")
            logger.debug(f"Initialization error traceback: {traceback.format_exc()}")
            raise

    def auto_connect(self):
        """Automatically connect to Outlook on startup"""
        logger.info("Auto-connect initiated")

        try:
            self.view.set_status("Connecting to Outlook...")

            def connect_thread():
                try:
                    logger.info("Connection thread started")
                    success, message = self.model.connect()
                    logger.info(f"Connection result: success={success}, message='{message}'")

                    # Update UI on main thread
                    self.view.after(0, lambda: self._on_connect_complete(success, message))

                except Exception as e:
                    logger.error(f"Error in connection thread: {e}")
                    logger.debug(f"Connection thread error traceback: {traceback.format_exc()}")
                    self.view.after(0, lambda: self._on_connect_complete(False, f"Connection failed: {str(e)}"))

            thread = threading.Thread(target=connect_thread, daemon=True)
            thread.start()
            logger.info("Connection thread launched")

        except Exception as e:
            logger.error(f"Error starting auto-connect: {e}")
            self.view.set_status("Error starting connection")
            self.view.show_error(
                "Startup Error",
                f"Failed to start connection process:\n{str(e)}"
            )

    def _on_connect_complete(self, success: bool, message: str):
        """Handle connection completion"""
        logger.info(f"Connection complete: success={success}")

        try:
            if success:
                self.view.set_status("Connected to Outlook - Click refresh to load emails")
                logger.info("Auto-loading conversations...")
                # Auto-refresh on successful connection
                self.refresh_conversations()
            else:
                logger.error(f"Connection failed: {message}")
                self.view.set_status("Failed to connect to Outlook")
                self.view.show_error(
                    "Connection Error",
                    f"{message}\n\nMake sure:\n"
                    "1. Microsoft Outlook is installed\n"
                    "2. Outlook is configured with an email account\n"
                    "3. You have pywin32 installed\n\n"
                    "Check outlook_reader.log for details."
                )

        except Exception as e:
            logger.error(f"Error in _on_connect_complete: {e}")
            logger.debug(f"Error traceback: {traceback.format_exc()}")

    def refresh_conversations(self):
        """Refresh conversations from Outlook"""
        logger.info("Refresh conversations requested")

        try:
            if not self.model.is_connected():
                logger.error("Cannot refresh - not connected to Outlook")
                self.view.show_error(
                    "Not Connected",
                    "Not connected to Outlook. Please restart the application."
                )
                return

            logger.info("Starting conversation refresh...")
            self.view.set_loading(True)
            self.view.clear_search()

            def load_thread():
                try:
                    logger.info("Load thread started")
                    conversations = self.model.get_conversations()

                    logger.info(f"Load thread completed: {len(conversations) if conversations else 0} conversations")

                    # Update UI on main thread
                    self.view.after(0, lambda: self._on_conversations_loaded(conversations))

                except OutlookDataError as e:
                    logger.error(f"OutlookDataError in load thread: {e}")
                    self.view.after(0, lambda: self._on_load_error(str(e)))

                except Exception as e:
                    logger.error(f"Unexpected error in load thread: {e}")
                    logger.debug(f"Load thread error traceback: {traceback.format_exc()}")
                    self.view.after(0, lambda: self._on_load_error(f"Unexpected error: {str(e)}"))

            thread = threading.Thread(target=load_thread, daemon=True)
            thread.start()
            logger.info("Load thread launched")

        except Exception as e:
            logger.error(f"Error starting refresh: {e}")
            logger.debug(f"Refresh error traceback: {traceback.format_exc()}")
            self.view.set_loading(False)
            self.view.set_status("Error starting refresh")
            self.view.show_error(
                "Refresh Error",
                f"Failed to start refresh:\n{str(e)}\n\nCheck outlook_reader.log for details."
            )

    def _on_conversations_loaded(self, conversations: List[Dict]):
        """Handle successful conversation loading"""
        logger.info(f"Processing loaded conversations: {len(conversations) if conversations else 0} items")

        try:
            # Validate conversations
            if conversations is None:
                logger.warning("Received None for conversations, using empty list")
                conversations = []

            if not isinstance(conversations, list):
                logger.error(f"Invalid conversations type: {type(conversations)}")
                conversations = []

            # Store conversations
            self.current_conversations = conversations
            logger.info(f"Stored {len(conversations)} conversations")

            # Display conversations
            logger.info("Displaying conversations in view...")
            self.view.display_conversations(conversations)

            # Update stats
            try:
                total_emails = sum(conv.get('count', 0) for conv in conversations if isinstance(conv, dict))
                self.view.update_stats(total_emails, len(conversations))
                logger.info(f"Updated stats: {total_emails} emails, {len(conversations)} conversations")
            except Exception as e:
                logger.error(f"Error updating stats: {e}")

            # Update status
            self.view.set_loading(False)

            if len(conversations) == 0:
                self.view.set_status("No emails found in inbox")
            else:
                total_emails = sum(conv.get('count', 0) for conv in conversations if isinstance(conv, dict))
                self.view.set_status(f"Loaded {len(conversations)} conversation(s) with {total_emails} email(s)")

            logger.info("Conversation loading completed successfully")

        except Exception as e:
            logger.error(f"Error in _on_conversations_loaded: {e}")
            logger.debug(f"Error traceback: {traceback.format_exc()}")
            self.view.set_loading(False)
            self.view.set_status("Error displaying conversations")
            self.view.show_error(
                "Display Error",
                f"Failed to display conversations:\n{str(e)}\n\nCheck outlook_reader.log for details."
            )

    def _on_load_error(self, error_message: str):
        """Handle conversation loading error"""
        logger.error(f"Conversation load error: {error_message}")

        try:
            self.view.set_loading(False)
            self.view.set_status("Error loading conversations")
            self.view.show_error(
                "Load Error",
                f"Failed to load conversations:\n{error_message}\n\nCheck outlook_reader.log for details."
            )
        except Exception as e:
            logger.error(f"Error in _on_load_error: {e}")

    def search_conversations(self, query: str):
        """Search conversations by query"""
        logger.info(f"Search requested: '{query}'")

        try:
            # Validate conversations list
            if not self.current_conversations:
                logger.info("No conversations to search")
                return

            if not isinstance(self.current_conversations, list):
                logger.error(f"Invalid current_conversations type: {type(self.current_conversations)}")
                return

            # Handle empty query
            if not query or not query.strip():
                logger.info("Empty query - showing all conversations")
                try:
                    self.view.display_conversations(self.current_conversations)
                    total_emails = sum(conv.get('count', 0) for conv in self.current_conversations if isinstance(conv, dict))
                    self.view.set_status(f"Showing all {len(self.current_conversations)} conversation(s)")
                except Exception as e:
                    logger.error(f"Error displaying all conversations: {e}")
                return

            # Filter conversations
            query_lower = query.lower().strip()
            filtered = []
            error_count = 0

            for conv in self.current_conversations:
                try:
                    if not isinstance(conv, dict):
                        logger.warning(f"Skipping invalid conversation in search")
                        error_count += 1
                        continue

                    # Search in subject
                    subject = conv.get('subject', '')
                    if isinstance(subject, str) and query_lower in subject.lower():
                        filtered.append(conv)
                        continue

                    # Search in sender names
                    emails = conv.get('emails', [])
                    if isinstance(emails, list):
                        for email in emails:
                            if isinstance(email, dict):
                                sender = email.get('sender', '')
                                if isinstance(sender, str) and query_lower in sender.lower():
                                    filtered.append(conv)
                                    break

                except Exception as e:
                    logger.error(f"Error filtering conversation: {e}")
                    error_count += 1
                    continue

            logger.info(f"Search found {len(filtered)} matching conversations ({error_count} errors)")

            # Display filtered results
            try:
                self.view.display_conversations(filtered)
                total_emails = sum(conv.get('count', 0) for conv in filtered if isinstance(conv, dict))
                self.view.set_status(f"Found {len(filtered)} conversation(s) matching '{query}'")
            except Exception as e:
                logger.error(f"Error displaying search results: {e}")
                self.view.set_status(f"Error displaying search results")

        except Exception as e:
            logger.error(f"Error in search_conversations: {e}")
            logger.debug(f"Search error traceback: {traceback.format_exc()}")
            self.view.set_status("Error performing search")

    def run(self):
        """Start the application"""
        logger.info("Starting main event loop...")
        try:
            self.view.mainloop()
            logger.info("Main event loop exited")
        except Exception as e:
            logger.error(f"Error in main event loop: {e}")
            logger.debug(f"Main loop error traceback: {traceback.format_exc()}")
            raise

    def cleanup(self):
        """Cleanup resources"""
        logger.info("Cleaning up resources...")
        try:
            self.model.disconnect()
            logger.info("Cleanup completed successfully")
        except Exception as e:
            logger.error(f"Error during cleanup: {e}")


def main():
    """Application entry point"""
    try:
        logger.info("Application starting...")
        app = OutlookInboxApp()
        app.run()
    except Exception as e:
        logger.error(f"Fatal application error: {e}")
        logger.debug(f"Fatal error traceback: {traceback.format_exc()}")
        print(f"\nFATAL ERROR: {e}")
        print(f"Check outlook_reader.log for details\n")
        input("Press Enter to exit...")
    finally:
        try:
            app.cleanup()
        except:
            pass
        logger.info("Application terminated")
        logger.info("="*60)


if __name__ == "__main__":
    main()
