"""
Main Window View - Modern GUI for Outlook Inbox Reader
Enhanced with comprehensive error handling
"""

import customtkinter as ctk
from typing import Callable, List, Dict, Optional
from datetime import datetime
import logging
import traceback

logger = logging.getLogger(__name__)


class MainWindow(ctk.CTk):
    """Main application window with modern design"""

    def __init__(self):
        super().__init__()

        # Configure window
        self.title("Outlook Inbox Reader")
        self.geometry("1200x800")

        # Set theme and color
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Configure grid layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Callbacks (to be set by controller)
        self.on_refresh_callback: Optional[Callable] = None
        self.on_search_callback: Optional[Callable] = None

        # Create UI components
        self._create_sidebar()
        self._create_main_area()
        self._create_status_bar()

    def _create_sidebar(self):
        """Create sidebar with controls"""
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.sidebar.grid_rowconfigure(6, weight=1)

        # App title
        self.logo_label = ctk.CTkLabel(
            self.sidebar,
            text="📧 Outlook Reader",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Subtitle
        self.subtitle_label = ctk.CTkLabel(
            self.sidebar,
            text="Inbox Conversations",
            font=ctk.CTkFont(size=14)
        )
        self.subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Search box
        self.search_label = ctk.CTkLabel(
            self.sidebar,
            text="Search:",
            font=ctk.CTkFont(size=13)
        )
        self.search_label.grid(row=2, column=0, padx=20, pady=(10, 5), sticky="w")

        self.search_entry = ctk.CTkEntry(
            self.sidebar,
            placeholder_text="Subject or sender..."
        )
        self.search_entry.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.search_entry.bind("<KeyRelease>", self._on_search_changed)

        # Refresh button
        self.refresh_button = ctk.CTkButton(
            self.sidebar,
            text="🔄 Refresh Inbox",
            command=self._on_refresh_clicked,
            font=ctk.CTkFont(size=14)
        )
        self.refresh_button.grid(row=4, column=0, padx=20, pady=10, sticky="ew")

        # Stats label
        self.stats_label = ctk.CTkLabel(
            self.sidebar,
            text="",
            font=ctk.CTkFont(size=12),
            justify="left"
        )
        self.stats_label.grid(row=5, column=0, padx=20, pady=10, sticky="w")

        # Appearance mode selector
        self.appearance_label = ctk.CTkLabel(
            self.sidebar,
            text="Appearance:",
            font=ctk.CTkFont(size=13)
        )
        self.appearance_label.grid(row=7, column=0, padx=20, pady=(10, 5), sticky="w")

        self.appearance_mode = ctk.CTkOptionMenu(
            self.sidebar,
            values=["Dark", "Light", "System"],
            command=self._change_appearance_mode
        )
        self.appearance_mode.grid(row=8, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.appearance_mode.set("Dark")

    def _create_main_area(self):
        """Create main content area"""
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Header
        self.header_label = ctk.CTkLabel(
            self.main_frame,
            text="Conversations",
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.header_label.grid(row=0, column=0, padx=0, pady=(0, 15), sticky="w")

        # Scrollable frame for conversations
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.main_frame,
            corner_radius=10
        )
        self.scrollable_frame.grid(row=1, column=0, sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        # Initial loading message
        self.loading_label = ctk.CTkLabel(
            self.scrollable_frame,
            text="Click 'Refresh Inbox' to load emails",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.loading_label.grid(row=0, column=0, pady=100)

    def _create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = ctk.CTkFrame(self, height=30, corner_radius=0)
        self.status_bar.grid(row=1, column=1, sticky="ew", padx=0, pady=0)

        self.status_label = ctk.CTkLabel(
            self.status_bar,
            text="Ready",
            font=ctk.CTkFont(size=11),
            anchor="w"
        )
        self.status_label.pack(side="left", padx=15, pady=5)

    def _on_refresh_clicked(self):
        """Handle refresh button click"""
        if self.on_refresh_callback:
            self.on_refresh_callback()

    def _on_search_changed(self, event=None):
        """Handle search text change"""
        if self.on_search_callback:
            query = self.search_entry.get()
            self.on_search_callback(query)

    def _change_appearance_mode(self, mode: str):
        """Change appearance mode (dark/light)"""
        ctk.set_appearance_mode(mode.lower())

    def set_status(self, message: str):
        """Update status bar message"""
        self.status_label.configure(text=message)

    def set_loading(self, is_loading: bool):
        """Show/hide loading state"""
        if is_loading:
            self.refresh_button.configure(state="disabled", text="Loading...")
            self.set_status("Loading conversations...")
        else:
            self.refresh_button.configure(state="normal", text="🔄 Refresh Inbox")

    def update_stats(self, total_emails: int, total_conversations: int):
        """Update statistics display"""
        try:
            stats_text = f"📊 Statistics\n\n"
            stats_text += f"Conversations: {total_conversations}\n"
            stats_text += f"Total Emails: {total_emails}"
            self.stats_label.configure(text=stats_text)
            logger.debug(f"Updated stats: {total_conversations} conversations, {total_emails} emails")
        except Exception as e:
            logger.error(f"Error updating stats: {e}")

    def display_conversations(self, conversations: List[Dict]):
        """Display conversations in the main area with comprehensive error handling"""
        logger.info(f"Displaying {len(conversations) if conversations else 0} conversations")

        try:
            # Validate input
            if conversations is None:
                logger.error("Received None for conversations list")
                conversations = []

            if not isinstance(conversations, list):
                logger.error(f"Invalid conversations type: {type(conversations)}")
                conversations = []

            # Clear existing content
            logger.debug("Clearing existing widgets...")
            try:
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
            except Exception as e:
                logger.error(f"Error clearing widgets: {e}")

            # Handle empty conversations
            if not conversations or len(conversations) == 0:
                logger.info("No conversations to display")
                try:
                    empty_label = ctk.CTkLabel(
                        self.scrollable_frame,
                        text="No conversations found",
                        font=ctk.CTkFont(size=14),
                        text_color="gray"
                    )
                    empty_label.grid(row=0, column=0, pady=100)
                except Exception as e:
                    logger.error(f"Error creating empty label: {e}")
                return

            # Display each conversation
            displayed_count = 0
            error_count = 0

            for i, conv in enumerate(conversations):
                try:
                    if not isinstance(conv, dict):
                        logger.warning(f"Skipping invalid conversation at index {i}: not a dict")
                        error_count += 1
                        continue

                    self._create_conversation_card(conv, i)
                    displayed_count += 1

                except Exception as e:
                    error_count += 1
                    logger.error(f"Error creating conversation card {i}: {e}")
                    logger.debug(f"Card creation error traceback: {traceback.format_exc()}")
                    continue

            logger.info(f"Displayed {displayed_count} conversations successfully ({error_count} errors)")

            # If all conversations failed to display, show error
            if displayed_count == 0 and len(conversations) > 0:
                logger.error("Failed to display any conversations")
                try:
                    error_label = ctk.CTkLabel(
                        self.scrollable_frame,
                        text="Error displaying conversations\nCheck log file for details",
                        font=ctk.CTkFont(size=14),
                        text_color="red"
                    )
                    error_label.grid(row=0, column=0, pady=100)
                except Exception as e:
                    logger.error(f"Error creating error label: {e}")

        except Exception as e:
            logger.error(f"Critical error in display_conversations: {e}")
            logger.debug(f"Display error traceback: {traceback.format_exc()}")
            try:
                error_label = ctk.CTkLabel(
                    self.scrollable_frame,
                    text="Critical error displaying conversations",
                    font=ctk.CTkFont(size=14),
                    text_color="red"
                )
                error_label.grid(row=0, column=0, pady=100)
            except:
                pass

    def _create_conversation_card(self, conv: Dict, row: int):
        """Create a card for a single conversation with error handling"""
        try:
            # Validate conversation data
            if not conv:
                logger.error("Empty conversation data")
                return

            # Get count with validation
            count = conv.get('count', 0)
            if not isinstance(count, int) or count < 0:
                logger.warning(f"Invalid count: {count}, defaulting to 0")
                count = 0

            # Determine if conversation has multiple emails
            is_multi = count > 1
            has_unread = conv.get('has_unread', False)

            # Get subject with validation
            subject = conv.get('subject', "(No Subject)")
            if not subject or not isinstance(subject, str):
                subject = "(No Subject)"

            # Create card frame
            try:
                # Build frame parameters
                frame_params = {
                    'master': self.scrollable_frame,
                    'corner_radius': 10,
                }

                # Add border only for unread conversations
                if has_unread:
                    frame_params['border_width'] = 2
                    frame_params['border_color'] = "#1f6aa5"

                card = ctk.CTkFrame(**frame_params)
                card.grid(row=row, column=0, sticky="ew", padx=5, pady=5)
                card.grid_columnconfigure(0, weight=1)
            except Exception as e:
                logger.error(f"Error creating card frame: {e}")
                return

            # Build subject text
            subject_text = subject
            if has_unread:
                subject_text = f"🔵 {subject_text}"
            if is_multi:
                subject_text = f"💬 {subject_text} ({count} messages)"

            # Create subject label
            try:
                subject_label = ctk.CTkLabel(
                    card,
                    text=subject_text,
                    font=ctk.CTkFont(size=15, weight="bold" if has_unread else "normal"),
                    anchor="w"
                )
                subject_label.grid(row=0, column=0, sticky="w", padx=15, pady=(12, 5))
            except Exception as e:
                logger.error(f"Error creating subject label: {e}")

            # Get emails list
            emails = conv.get('emails', [])
            if not isinstance(emails, list):
                logger.warning(f"Invalid emails type: {type(emails)}")
                emails = []

            # Display emails in conversation
            if emails:
                # Show last 3 emails
                recent_emails = emails[-3:] if len(emails) > 3 else emails
                for j, email in enumerate(recent_emails):
                    try:
                        if isinstance(email, dict):
                            self._create_email_row(card, email, j + 1, len(emails) > 3 and j == 0)
                        else:
                            logger.warning(f"Skipping invalid email at index {j}")
                    except Exception as e:
                        logger.error(f"Error creating email row {j}: {e}")

            # If more than 3 emails, show indicator
            if count > 3:
                try:
                    more_label = ctk.CTkLabel(
                        card,
                        text=f"... and {count - 3} more message(s)",
                        font=ctk.CTkFont(size=11),
                        text_color="gray"
                    )
                    more_label.grid(row=100, column=0, sticky="w", padx=15, pady=(0, 10))
                except Exception as e:
                    logger.error(f"Error creating 'more' label: {e}")

        except Exception as e:
            logger.error(f"Error in _create_conversation_card: {e}")
            logger.debug(f"Conversation card error traceback: {traceback.format_exc()}")

    def _create_email_row(self, parent, email: Dict, row: int, show_ellipsis: bool):
        """Create a row for a single email within a conversation with error handling"""
        try:
            # Validate email data
            if not email or not isinstance(email, dict):
                logger.warning("Invalid email data")
                return

            # Format timestamp safely
            received_time = email.get('received_time')
            try:
                if received_time and hasattr(received_time, 'strftime'):
                    time_str = received_time.strftime("%b %d, %Y %I:%M %p")
                else:
                    time_str = "Unknown date"
            except Exception as e:
                logger.warning(f"Error formatting timestamp: {e}")
                time_str = "Unknown date"

            # Get sender safely
            sender = email.get('sender', 'Unknown')
            if not sender or not isinstance(sender, str):
                sender = 'Unknown'

            # Get unread status
            unread = email.get('unread', False)

            # Create info text
            unread_indicator = "🔵 " if unread else ""
            info_text = f"{unread_indicator}{sender} • {time_str}"

            # Create label
            try:
                info_label = ctk.CTkLabel(
                    parent,
                    text=info_text,
                    font=ctk.CTkFont(size=12),
                    anchor="w",
                    text_color="lightblue" if unread else "gray"
                )
                info_label.grid(row=row, column=0, sticky="w", padx=35, pady=2)
            except Exception as e:
                logger.error(f"Error creating email info label: {e}")

        except Exception as e:
            logger.error(f"Error in _create_email_row: {e}")
            logger.debug(f"Email row error traceback: {traceback.format_exc()}")

    def show_error(self, title: str, message: str):
        """Show error dialog"""
        error_window = ctk.CTkToplevel(self)
        error_window.title(title)
        error_window.geometry("500x200")
        error_window.transient(self)
        error_window.grab_set()

        # Center the window
        error_window.update_idletasks()
        x = (error_window.winfo_screenwidth() // 2) - (500 // 2)
        y = (error_window.winfo_screenheight() // 2) - (200 // 2)
        error_window.geometry(f"+{x}+{y}")

        # Error icon and message
        icon_label = ctk.CTkLabel(
            error_window,
            text="⚠️",
            font=ctk.CTkFont(size=48)
        )
        icon_label.pack(pady=(20, 10))

        msg_label = ctk.CTkLabel(
            error_window,
            text=message,
            font=ctk.CTkFont(size=13),
            wraplength=450
        )
        msg_label.pack(pady=10, padx=20)

        # OK button
        ok_button = ctk.CTkButton(
            error_window,
            text="OK",
            command=error_window.destroy,
            width=100
        )
        ok_button.pack(pady=20)

    def clear_search(self):
        """Clear search box"""
        self.search_entry.delete(0, "end")
