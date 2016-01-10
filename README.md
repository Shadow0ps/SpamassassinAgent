Spamassassin Agent for Microsoft Exchange Server
==================

> To work correctly you must build the solution with references to 
> the correct Microsoft Exchange DLLs.

This transport Agent Interfaces Microsoft Exchange directly to SpamAssassin.

Spamassassin runs in two parts, the daemon (spamd) and the client (spamc). This plugin will take a message, feed it to the Spamassassin client. The spamassassin client will then connect to the Spamassassin Server and run the spamassassin scoring software on the message. Once the message is returned, the agent will then attempt to find the score
if the score is above the discard threshold the agent will tag the message with X-Spam-Discard: YES. Then the Exchange Server can take action based on this header. 


Features
-----
1. Configurable Max Message size. This avoids memory issues when scanning very large messages
2. Configurable Discard threshold (10+)
3. Bayesian learning support 
4. Full logging library


Install Instructions:
-----
https://kill-9.me/557/exchange-spamassassin-transport-agent


Contact
-----
- James DeVincentis
- james@hexhost.net
