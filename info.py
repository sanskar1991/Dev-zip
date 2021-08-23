"""
I had a question and a suggestion for the 'unpack' func that:
We are passing the 'input_deck's' name without specifing the path as 'fqfn' in 'unpack' func
that means let's take an exapmle of Onboarding

In unpack func, 
fqfn  = Onboarding
fp = /tmp/{fqfn} or /tmp/Onboarding


In order to extract the input_deck (Onboarding.pptx) we need to give full system path of it and
the full system path of the output directory
i.e., 'fp' and 'fqfn' should contains the full system path
"""

"""
def contents: 

what I'm doing is taking input path, output path and a dict of refactored_name as arguments

then just getting roots of [Content_Type].xml file
root1: root of input_deck
root2: root of output_deck

collecting all the content types if relation's tag is 'Default'

then iterating all the relations in root1 (as in input_deck's [Content-Tyoe].xml file)
and then checking if the asset is present in the refactored_name dict (means in it's keys)
then I'm refactoring the name of that asset and changing it in the relationship itself

But what we are trying to do is getting all the relations of the assets that needs to be copied
and then create a dict

so what I did is I just added the asset name as key and the relationsip object as its value
so that we can just refactor it based on the asset name (keys) and simply change it in its relationship
and then add it in the output_deck's [Content_Types].xml file

"""
########| Day 2 |########

"""
A quick question, as you said that you wanted to have a dictionary with all the asset names as keys 
and the whole file, i.e., the content of the file as value

so key would be only asset name or asset_name with full path?

"""

"""
Changes:
1st:
In 'process_message' func, here it is written "ss = msg.get('s', None)" right, so let's take an example
if we have slides as [2, 4, 6] then ss will be a list but if "None" passed then we are just taking 
"ss" as an integer. So, I changed it to list from 1 to total number.

2nd:
Added a condition in "assetfn_to_relfn" to only process those targets which have '/' in it
otherwise return "file_name.rels" (pause for 2 sec). So, when we check for its path then it will return "False"

3rd:
As you said that you wanted to have a dictionary with all the asset names as keys and the whole file,
i.e., the content of the file as value
So, I did this by using "lxml" library of python. I added the asset name as key and its root and tree
as value so that we can chan

4th:
In my approach I was first creating an empty deck, iterating over the message and one by one merging
the input deck to output deck.. now what we are doing is we are collecting all the assets of all the 
input deck into a dictionary "assets" and then merge them.
"""

########| Day 3 |########

"""
I'm done with all the "build" funcs as in build_assets, rels, content_types and properties. 
Also I'm done with the apply_properties and apply_contents fucntions.
In doing so I got few things which I wanted to discuss before moving further (pause)

1st:
When we create an empty deck and unpack it we get:
o: 1 slideMasters
o: 11 slideLayouts
o: printerSetting files
These files are of no use to us as we'll start new asset name from slidemaster1, slideMaster2, so on
and for slideLayouts it will be slideLayout1, slideLayout2 and so on..
So, I handled it by removing those files before start adding assets of the first input_deck

I think we should do this here as well just after unzipping the empty deck (pause)
So, what should we do?

2nd:
For properties files like ('tableStyles.xml', 'commentAuthors.xml', 'presProps.xml') which I was 
handling differently in my code, by comparing input deck's property file with output deck's property file to see 
if any field is missing in any tag or some figures are different. After comparing if I find anything 
then I just add that to the output deck's property file and save it.

So if we'll do same here then we need only one function instead of build_properties and apply_properties
functions. If not then we can simply make a dict of all the input deck's property files and 
their contents, pass that dict to apply_property func, perform comparision and then modify 
existing property files of output deck.

I implemented both here so we can decide which way we'll be going.

3rd:
For Content_Types.xml file which contains all the content type of assets, we need to modify existing 
file of output deck's. We need to add content_type of new assets with refactored names. That means
we first need to refactor the names in the 'contents' dict itself and then add those content types 
in the output deck's content_types.xml file

So, I have written the code for just adding them without refactoring asset names first.
I wanted to discuss this first, if you have a something in mind

Because.. (continue with 4th point)

4th:
What I was doing for assets, rel files, content_types, I was first changing their names means
refactoring their names with the help of a dict "refactored_count" and then add those files 
in the output deck. So I wanted to ask that how are we gonna handle the refactoring of asset names?

Also, we need to change these refactored names in then content of all the rels files. 

So, what I think is we need to create a dict with old asset names as keys and refactored names as value
which we refer before creating assets in the output deck.

So, in other words.. (continue with 5th point)

5th:
I wants to discuss about the refactoring of asset names and rIds.
For doing this I was creating a dict "refactored_names" consists of keys as old asset names and
value as "refactored name" as per the number of files present in a directory and then copying 
all the assets and pasting them with their new refactored name

But now we don't have any dictionary to track number of assets present in a dir so what do you suggest
for this

What I suggest is we should create a dict, refactored_assets to track the count of the assets like..
let me show you (I'll navigate to the file)...
Here as you can see we are counting number of assets of each type so that if any new asset comes 
we can use this value to rename it.
For example, if a new slide will come we can simply incremnt this value and assign new name 
based on the incremented count

So.. yes

"""

########| Day 4 |########

"""
QUE: 
ctx['asstes'] will have a dict of all the assets with their connect
ctx['rels'] will have a dict of all the asset's rel files with their connect

I created two dicts:
1. refactored_nm (name)
2. refactored_cnt (count)

refactored_nm (name) stores all old name of an asset as key and refactored_name of assest as value
refactored_cnt (count) stores all the asset names as keys and their count as value

Now, the next step is to update all the rel file's content with refactored names of assets
for example, if a rel file have a media file named "image10.png" which is renamed as "image5.png"
then we have to update that old name with new name.

I am working on this

"""

##### DAY XXXX WED #####

"""
I worked on the refactoring of contents of rel files in which I used the refactored_fns (refactored filenames)
dict to get the asset names and modified the root of the rel files

So, for doing this I made few changes in the refactored_fns(file names) dict (I'll navigate)
Here as you can see I changed the new name to just last two elements of 'L'
it is because at the end we'll have to remove the previous path so I handled it here only.

then I created a function "refactor_content" which firstly I extracted the input_deck name with full path
as 'ft' then here I'm iterating the root and generating "new_name" for every asset
then just setting the "new_name"

Here I'm not returning anythng, if ypu want then we can return something but we don't need here

After this I'm just calling our "refactor_assets_and_rels" func which will rename then rel file and
save it in the output deck. 

2nd:
Till now we were handling all the assets other than mandatory assets as we wreated a func "is_mandatory"
which return True if the passed asset does not belongs to slideMasters, slideLayouts, themes.
So in order to handle the slideMasters, themes and Layouts I changed the name of func "is_mandatory"
to "is_required" and "build_mandatory_assets" to "build_required_assets"
and now created separate function for the mandatory assets named "build_mandatory_assets" and 
"apply_mandatory_assets".

I completed them but I think we can figure out some other way to handle them.


3rd: (only when he asks) Why we can't include these mandatory assets in out normal assets?
Ans.: It is because we don't need to walk asset tree for these files, we just need to get these files
rename them and then apply them in the output deck I did this in my code and it was working
But here we can merge this with assets but I'm not sure.

If he asks WHY YOU ARE UNSURE?
Ans.: There is no reason for my unsurity, it's just I haven't tried it.. that's it


4th:
Consern:
I have one consern about walking asset tree and adding all the "target" in the asset's dict.
It will work fine for all the slides but it we might get into an infinite loop

for example: handoutMaster1.xml.rels we have "slide6.xml" in the target of its rel files
so when we walk asset tree for its rel file we'll again go to "slide6.xml's" rel file
and will do the walking in a loop.

So, what I think is we need to add a condition which will not allow to walk asset tree for slide
if we already have it in the asset dict
or need to figure out some other way to handle this 

"""

"""
I completed the "apply_content_types" func as it required the "refactored_file_names" dict so
you can have a look it.. (I'll navigate)
Also, I created a function "modify_output_xml" to handle the modification part which I used in the 
"apply_properties" func as well

One more thing, did you remember we created a func "remove_default_files" that removes the default files
like slideMasters, Layouts and themes. So if I talk about this "Content_Type.xml" file then it also
contains content types of those files which we don't want.
To remove it we have 2 ways:
1. Either we can handle this in "remove_default_files" func by opening the "Content_Types.xml" file
and iterate all the relation to check for these default files and remove them
2. Or we can do same thing in "apply_contents" func after populate content types of all
the assets by iterating the file and removing duplicate relations

I can implement both so what do you suggest?

"""

"""
To implement "refactoring of rId" I'm thinking of creating a new func which deals with both
presentation.xml.rels as well as presentation.xml because we need to de refacroting of rIds in 
those files

So, I created a func "build_pxr_file" for "presentation.xml.rels" in which I'm collection all the relations other than slides which
are not mentioned in the "ss" as in we are getting all the assets and slides relation which are required.
I did this.. and for "apply_pxr" func, I will store all the relations of the pxr file of output deck's and then will compare and 
remove duplicate entries from the dictionary. Then I'll use the refactored_filenames dict to refactor the names and 
refactor r_Ids by using "lasts" dictionary and then save it.
So this is what I'm thinking to do..
"""

"""
Flow of the code:
1. Gather all the required assets and relation from all the input decks one by one
a. build/gather all the assets and their rel files with content and store in ctx
b. similarly gather all the content types of all the required assets and store in ctx
c. gather all the property files and store in ctx
d. gather the presentation files, that are presentation.xml.rels and presentation.xml and store in ctx
2. Apply them in the output deck
a. refactor asset names and create them
b. refactor rel files's content and their names and create them
c. refactor asset names in relation and add them in the Content_types.xml file
d. add the properties all the property files and if a property file is not exist then create it
e. refactor asset names and their rIds in relations then add them in the presentation.xml.rel(pxr) file
f. refactor the rIds of the assets and then add them in the presentation.xml file
3. Zip the output deck directory
"""
"""
Remaining part:
As per the points mentioned above, 
0- 1-a: might need some modifiction
0- 2-e: refactoring of rId
0- 2-f: whole
"""