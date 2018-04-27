# ATLAS
AuTomated LAbbook System

Electronic system for managing lab notes, lab organisation, etc.

Written in Python3, visualised by PyQt5.
Uses sqlite3 for database storage
Uses docx for output of information in document format

Compartmentalised Lab Notes
---------------------------

ATLAS aims to allow rapid and detailed experimental notes. 

Experiments are organised by alphabetical categories and by number (e.g. Experiment F3 is the third experiment from Project F). Each Experiment has its own directory with a Word Document at its core (Experiments/F Experiments/Experiment F3/Experiment F3.docx). Experimental information is entered into the ATLAS GUI and saved in document format.

The Core Document begins with a title, starting date, rationale and short results summary (filled when experiment is complete). After this, the document is appended with Protocols, Updates or Links. 

Protocols are pre-written experimental plans with numerous variables throughout. Protcols are initially written in the ATLAS GUI and can be edited here too. The user is able to quickly scan through and alter variables. When finished, the full protocol is inserted into the Core Document with the variables in place. e.g.


"x ml Super Optimal Broth was added to the cells and they were then incubated at y oC for z minutes." 

Default variables will be based on most common usage and may read:

"0.2 ml Super Optimal Broth was added to the cells and they were then incubate at 37 oC for 60 minutes."

User can move through the text variable-by-variable and alter each one using a drop-down menu or by entering new text.


Updates are short observations which do not fit into pre-written Protocol formats. 

Links are Document-style hyperlinks to data which is not suited to the Document format but which is stored in the same folder. Links are also automatically generated when another Experiment is referenced. 


The Experiment organisation is centred on the ubiquitous Document type to allow easy sharing between lab members using different computers and systems. The Share option will zip an Experiment folder and place a copy on the Desktop. 


Timed Lab Notes
---------------

Protocols, Updates and Links are also collected in a separate Document file as they are created. When physical lab books are still required, this file can be printed to give a list of lab progress by date.


Plasmids
--------

Details of plasmids, primers and other lab materials can be organised using a database. The default database for plasmids allows entry of the Name, Source, Location, Putative Sequence, Known Sequence Information. 

TBC.




