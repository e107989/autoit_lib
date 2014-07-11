#include "utilities.au3"

$test_str = "abcd123"&@CRLF&"efgh456"&@CRLF&"ijkl789"&@CRLF&"mnop101"
AlertArray(searchForInWizardScreen($test_str, "op", 2, 1, 4, 3))