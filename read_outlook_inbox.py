"""
Microsoft Outlook Inbox Reader
Reads and displays email subjects from Outlook Inbox grouped by conversation
"""

import win32com.client
from collections import defaultdict
from datetime import datetime


def read_outlook_inbox_conversations():
    """
    Connect to Microsoft Outlook via COM and read inbox emails grouped by conversation.
    Returns a dictionary of conversations.
    """
    try:
        # Connect to Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Access the Inbox folder (folder index 6 is Inbox)
        inbox = namespace.GetDefaultFolder(6)

        # Get all messages from inbox
        messages = inbox.Items

        # Group messages by conversation
        conversations = defaultdict(list)

        for message in messages:
            try:
                # Get conversation ID and email details
                conv_id = message.ConversationID if hasattr(message, 'ConversationID') else None
                subject = message.Subject if message.Subject else "(No Subject)"
                received_time = message.ReceivedTime if hasattr(message, 'ReceivedTime') else None
                sender = message.SenderName if hasattr(message, 'SenderName') else "Unknown"

                # Use conversation ID or subject as grouping key
                # If no conversation ID, use subject as fallback
                group_key = conv_id if conv_id else f"single_{subject}"

                conversations[group_key].append({
                    'subject': subject,
                    'sender': sender,
                    'received_time': received_time,
                    'conv_id': conv_id
                })

            except Exception as e:
                # Skip messages that cause errors
                print(f"Warning: Could not read message - {str(e)}")

        # Sort each conversation by received time (oldest first within conversation)
        for conv_id in conversations:
            conversations[conv_id].sort(
                key=lambda x: x['received_time'] if x['received_time'] else datetime.min,
                reverse=False
            )

        # Convert to list and sort conversations by most recent message (newest first)
        conversation_list = []
        for conv_id, emails in conversations.items():
            latest_time = max(
                (email['received_time'] for email in emails if email['received_time']),
                default=datetime.min
            )
            conversation_list.append({
                'conv_id': conv_id,
                'emails': emails,
                'latest_time': latest_time,
                'count': len(emails)
            })

        conversation_list.sort(key=lambda x: x['latest_time'], reverse=True)

        return conversation_list

    except Exception as e:
        print(f"Error accessing Outlook: {e}")
        print("\nMake sure:")
        print("1. Microsoft Outlook is installed")
        print("2. You have pywin32 installed (pip install pywin32)")
        print("3. Outlook is configured with an email account")
        return []


def main():
    """Main function to display inbox emails grouped by conversation"""
    print("Reading Outlook Inbox (Grouped by Conversation)...\n")

    conversations = read_outlook_inbox_conversations()

    if conversations:
        total_emails = sum(conv['count'] for conv in conversations)
        print(f"Found {total_emails} email(s) in {len(conversations)} conversation(s):\n")
        print("=" * 100)

        for i, conversation in enumerate(conversations, 1):
            email_count = conversation['count']

            # Display conversation header
            if email_count > 1:
                print(f"\n[Conversation {i}] - {email_count} emails")
            else:
                print(f"\n[Email {i}]")

            print("-" * 100)

            # Display each email in the conversation
            for j, email in enumerate(conversation['emails'], 1):
                received_str = email['received_time'].strftime("%Y-%m-%d %H:%M") if email['received_time'] else "Unknown"

                if email_count > 1:
                    print(f"  {j}. [{received_str}] {email['sender']}: {email['subject']}")
                else:
                    print(f"  [{received_str}] {email['sender']}: {email['subject']}")

        print("\n" + "=" * 100)
    else:
        print("No emails found or unable to access Inbox.")


if __name__ == "__main__":
    main()
