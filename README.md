# SSRSAggLookup
VB code for embedding in an SSRS report, allows aggregating in a number of ways over a lookup set

To use paste into the Code window of the Report Properties, then use the following expression in an object:

=code.AggLookup([AggregateType as String 1], [Lookupset 2])

1: Must be lower case, can be sum, count, avg, min, max

2: Use SSRS LookupSet expression, i.e. LookupSet([LocalMatch], [TargetMatch], [ReturnValue], [Dataset as String])


