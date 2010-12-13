==================================
Caching DNS Server in Visual Basic
==================================

.. image:: https://github.com/jparise/caching-dns/raw/master/screenshot.png
   :align: right

Years ago, I was enrolled in the `RIT Information Technology`_ department's
Network Programming course.  Because the "official" language of the IT
department at the time was Visual Basic, that was the required programming
language for all of the course's projects.  I really can't say I was a big fan
of Visual Basic, but it got the job done (even if the job is network
programming ...).

One of our projects dealt with the Domain Name Service protocol (as defined by
`RFC 1035`_).  We were tasked with writing a caching name server that would
forward unknown requests to a configurable parent name server.  The cache has
to honor each entry's expiration (TTL).

The most difficult part of the project was crunching the internal DNS packet
format.  The original creators of the packet format implemented a "pointer"
method, presumably to save space by reducing redundant strings inside of the
packet.  For example, if the string ``example.com`` appears in the packet
before a string that is intended to represent ``host.example.com``, the second
string is shortened to ``host.offset``, where *offset* is a numeric offset
from the start of the packet to the first occurence of the ``example.com``
string.

There is no limit to the number of these pointers that can be used in a given
packet, so long as they are used correctly.  It took me a while to completely
grok this problem, but I solved it somewhat cleanly using a recursive
function.  I don't think I had it working perfectly in all cases, however, but
it was good enough to function correctly under normal use.

You're perfectly free to use this code for anything you like, provided you
conform to the terms of the standard `BSD License`_ (the text of which is
included in each of the source files).

I don't offer any support for this software nor do I intend to continue its
development.  I provide it here as a fairly good example of sockets
programming under Windows using Visual Basic.  It's more of a learning toy
than anything else, and you should treat it as such.

Enjoy!

.. _RIT Information Technology: http://www.ist.rit.edu/
.. _RFC 1035: http://www.ietf.org/rfc/rfc1035.txt
.. _BSD License: http://www.opensource.org/licenses/bsd-license.php
