#include <pebble.h>

char* getSubString(char* str, int startIndex, int endIndex)
{
    char* newStr = (char*)malloc(endIndex - startIndex + 1);
    int count = 0;
    for(int i = startIndex; i<endIndex; i++)
    {
        newStr[count] = str[i];
        count++;
    }
    newStr[count] = '\0';
    return newStr;
}

void outputArrayContents(char** array, int arrayLength)
{
    for(int i=0; i<arrayLength; i++)
    {
        APP_LOG(APP_LOG_LEVEL_DEBUG, "\nArray Element %d: %s\n", i, array[i]);
    }
}
//return an array of items which have been split by delimiter
char** splitString(char* str, char delimiter, int* length)
{
    char** splitArray = (char**)malloc(sizeof(char*));
    int previousDelimiterIndex = 0;
    int tokenCount = 0;
            
    for(int i=0; i<(int)strlen(str); i++)
    {
        //if we hit the delimiter or end of string execute below code
        if ((str[i] == delimiter) || (i == ((int)strlen(str)-1)))
        {
            tokenCount++;
            splitArray = (char**)realloc(splitArray, sizeof(char*)*(tokenCount));
            splitArray[tokenCount-1] = (char*)malloc((i-previousDelimiterIndex) + 1);
            //memset(splitArray[tokenCount-1], '\0', (i-previousDelimiterIndex) + 1);
            //(i-previousDelimiterIndex) + 1
            char* derp = getSubString(str, previousDelimiterIndex, i);
            splitArray[tokenCount-1] = strcpy(splitArray[tokenCount-1], derp);
            //outputArrayContents(splitArray, tokenCount);
            previousDelimiterIndex = i+1;
        }
    }
    *length = tokenCount;
    
    return splitArray;
}

void scrollTextByOneLetter(char* text)
{
    int length = strlen(text);
    char tmp = text[0];
    for(int i=0;i < (length-1);i++)
    {
        text[i] = text[i+1];
    }
    text[length-1] = tmp;
}

//this will copy a substring from src to dst and if it hits the end of the string it will loop
//back around and start from the beginning of string.
void loopingCopy(char* dst, char* src, int startIndex, int length)
{
    int i = 0;
    for(i=0;i < length; i++)
    {
        dst[i] = src[startIndex];
        if ((startIndex) == ((int)strlen(src) - 1))
        {
            startIndex = 0;
            continue;
        }
        startIndex++;
    }
    //dst[i + 1] = '\0';
}