# GatherMobileSIMData

A problem was positioned with me many years ago to gather Mobile SIM data for any devices that had a sim in the machine whether the sim was active from the provider or just prepopulated waiting for the customer to activate the SIM.
So I came up with this.  it also attaches nicely into SCCM for inventory.

The system basically looks at the netsh data available and then gathers the readystate information for the connections.  The Reeadystate is where the SIM ID sort of lives.....  Its not a simple read data and there you go.
The script also converts the SIMID and provides a checksum digit for the final piece of the numbers required for the providers to connect the sim.

To install into the hardware inventory, simply run the gather script on one device and then add the hardware inventory via the browse method.  I will list out how to do this explicitly later.

To be continued
