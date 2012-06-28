import uno
import re
from copy import copy

#def log(message):
#    f = file("C:\Apps\OOlog.txt","a+")
#    f.write(message + "\n")
#    f.close()

class RunScript(object):
    """
    This is a convenience class for calling external scripts using natural
    Python function calling syntax.
    """
    def __init__(self, pkg, language, location):
        self.pkg      = pkg
        self.language = language
        self.location = location
        # TODO: Decide if we want to keep last result.
        #self.last_result = None

    def __getattr__(self, name):
        def call(*args,**kw):
            pkg      = kw.get("pkg",self.pkg)
            language = kw.get("language",self.language)
            location = kw.get("location",self.location)
            if   language == "Python": script = pkg + "$" + name
            elif language == "Basic":  script = pkg + "." + name
            else: script = pkg + name

            ctx = uno.getComponentContext()
            scripts = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
            scriptProvider = scripts.createScriptProvider(" ")
            url = "vnd.sun.star.script:%(script)s?language=%(language)s&location=%(location)s" % {'script':script,'language':language,'location':location}
            myScript = scriptProvider.getScript(url)
            if args is None: args = ()
            result = myScript.invoke(args,(),())
            #self.last_result = result
            # TODO: Decide whether there is a case where this is not a tuple or where values a that are not the first value matter.
            return result[0]

        return call

OOTools = RunScript("OOTools.py","Python","share")

class PyInputField(object):
    def __init__(self, field):
        self.field = field

    def __getattr__(self, name):
        if name == "Name":
            return self.field.Content
        return getattr(self.field, name)

class PyUserField(object):
    def __init__(self, field):
        self.field = field

    def __getattr__(self, name):
        if name == "Name":
            return self.field.TextFieldMaster.Name
        return getattr(self.field, name)

def PyTextFieldFactory(field):
    """
    This is a way to make Inputfields and UserField behave in the same way.
    We want to make sure that attribute access is the same for both. This is
    a case of the Adapter pattern.
    """
    if field.supportsService("com.sun.star.text.textfield.InputUser"):
        return PyInputField(field)

    if field.supportsService("com.sun.star.text.textfield.User"):
        return PyUserField(field)

    return None

def GetBlockTextField(XSCRIPTCONTEXT, block):
    doc = XSCRIPTCONTEXT.getDocument()
    fields = doc.getTextFields()
    enum   = fields.createEnumeration()
    block_field = None
    while enum.hasMoreElements():
        field = PyTextFieldFactory(enum.nextElement())
        if field == None: continue
        if field.Name == block:
            block_field = field
    return block_field

class BaseNameHandler(object):

    FIRST_IN_LIST = object()

    def __init__(self):
        self.listeners = []

    def AddUpdateListener(self,listener):
        """
        Listeners are a function that recieves the NameHandler as the first
        parameter, and a list of added and removed items since the last update.
        Additional named parameters may be sent as well.
        """
        if listener not in self.listeners:
            self.listeners.append(listener)

    def DropUpdateListener(self,listener):
        if listener in self.listeners: self.listeners.remove(listener)

    def __CallListeners__(self, *args, **kwargs):
        for listener in self.listeners:
            listener(*args,**kwargs)

    def GetPreviousName(self, name):
        """
        Get the name before the one given. If None in returned, the name did
        not exist in the list. If FIRST_IN_LIST is returned, the name was the
        first in the list.
        """
        pass

    def GetNames(self, XSCRIPTCONTEXT=None):
        pass

    def Update(self, XSCRIPTCONTEXT):
        self.GetNames(XSCRIPTCONTEXT)

class NameHandler(BaseNameHandler):
    def __init__(self, prefix):
        BaseNameHandler.__init__(self)
        self.names = []
        self.prefix = prefix
        self.listeners = []

    def GetPreviousName(self,name):
        if name not in self.names: return None
        else:
            index = self.names.index(name)
            if index == 0: return self.FIRST_IN_LIST
            return self.names[index-1]

    def GetPrefix(self):
        return self.prefix

    def GetNewName(self, XSCRIPTCONTEXT):
        base = self.GetPrefix()
        prefix = r"com\.sun\.star\.text\.fieldmaster\.User\.%s(\d+)" % base
        check = re.compile(prefix)
        names = XSCRIPTCONTEXT.getDocument().getTextFieldMasters().getElementNames()
        numbers = [0]
        for name in names:
            match = check.match(name)
            if match:
                #OOTools.MsgBox(match.group(1))
                num = int(match.group(1))
                numbers.append(num)
        name = "%s%d" % (base,max(numbers)+1)
        return name

    def GetNames(self, XSCRIPTCONTEXT=None):
        if XSCRIPTCONTEXT is None: return self.names
        prefix = self.GetPrefix()

        # TODO: check that XSCRIPTCONTEXT is a writer doc.
        model  = XSCRIPTCONTEXT.getDocument()
        fields = model.getTextFields()
        enum   = fields.createEnumeration()
        # First get the list of fields in order
        textfields = []
        names = []
        while enum.hasMoreElements():
            try:
                field  = PyTextFieldFactory(enum.nextElement())
                if field is None: continue
                name   = field.Name
                if name.find(prefix) != -1:
                    #log("Found: " + name)
                    added = False
                    # Put the fields in the array in order
                    for textfield in textfields:
                        if model.Text.compareRegionStarts(field.Anchor, textfield.Anchor) == 1:
                            index = textfields.index(textfield)
                            names.insert(index, name)
                            textfields.insert(index, field)
                            added = True
                            #log("Inserted: " + name)
                            break
                    # If we need to add this at the end of the array.
                    if not added:
                        names.append(name)
                        textfields.append(field)
            except:
                pass

        # Compare the names to the stored set of names. If there is a change, notify listeners.
        call_listeners = False
        if len(self.names) != len(names): call_listeners = True
        else:
            pairs = zip(self.names,names)
            for a,b in pairs:
                if a != b:
                    call_listeners = True
                    break

        if call_listeners:
            added   = []
            removed = copy(self.names)
            self.names = names
            for name in names:
                if name in removed:
                    removed.remove(name)
                else:
                    added.append(name)

            self.__CallListeners__(self,added=added,removed=removed)
        return names

    def Update(self, XSCRIPTCONTEXT):
        self.GetNames(XSCRIPTCONTEXT)

class SequencedNameHandlers(BaseNameHandler):
    def __init__(self, name_handlers):
        BaseNameHandler.__init__(self)
        #OOTools.MsgBox("Got Here: " + str(name_handlers))
        self.name_handlers = name_handlers

        for handler in self.name_handlers:
            handler.AddUpdateListener(self.UpdateListener)

        self.hold_update_mode = False
        self.added   = []
        self.removed = []

    def UpdateListener(self, handler, added=[], removed=[]):
        self.added.extend(added)
        self.removed.extend(removed)
        if not self.hold_update_mode:
            self.__CallListeners__(self,added=self.added,removed=self.removed)
            self.added   = []
            self.removed = []

    def GetPreviousName(self, name):
        prev = None
        for handler in self.name_handlers:
            #OOTools.MsgBox(str(len(self.name_handlers)))

            name = handler.GetPreviousName(name)
            if name == self.FIRST_IN_LIST:
                if prev == None:
                    return self.FIRST_IN_LIST
                else:
                    return prev.GetNames()[-1]
            else:
                return name
            if len(hanlder.GetNames()) > 0:
                prev = handler
        return None

    def GetNames(self, XSCRIPTCONTEXT=None):
        names = []
        for handler in self.name_handlers:
            names.extend(handler.GetNames(XSCRIPTCONTEXT))

        return names

    def Update(self, XSCRIPTCONTEXT):
        self.hold_update_mode = True
        try:
            for handler in self.name_handlers:
                try:
                    handler.Update(XSCRIPTCONTEXT)
                except:
                    pass
        finally:
            self.hold_update_mode = False
        self.UpdateListener(None)

class NameHandlerFactoryImpl(object):
    def __init__(self):
        self.name_handlers = {}
        self.types = ['Local','Global','Master']

    def __GetTypeBaseString__(self,typ):
        if   typ == "Local":  return "CAS_L"
        elif typ == "Global": return "CAS_G"

    def GetTypes(self):
        return copy(self.types)

    def GetNameHandler(self, typ):
        if typ not in self.types: raise KeyError("Type %s not found in NameHandlerFactory") % typ
        if typ not in self.name_handlers:
            if typ == "Master":
                handlers = (self.GetNameHandler("Global"),self.GetNameHandler("Local"))
                name_handler = SequencedNameHandlers(handlers)
            else:
                name_handler = NameHandler(self.__GetTypeBaseString__(typ))
            #OOTools.MsgBox(typ)
            self.name_handlers[typ] = name_handler
        return self.name_handlers[typ]

    def __call__(self, typ):
        return self.GetNameHandler(typ)

class DictLinkWrapper(object):
    """
    This is used to link together dictionaries. It requires a wrapped dictionary and a parent
    dictionary. Items are looked up in the wrapped dictionary first, then in the parent
    dictionary.

    This is intended to be used as a way to store the local stack when executing code blocks
    inside a writer document. When exec is called, a global and local stack can be passed in.
    Each block should have its own local 'prestack' that comes from a linked together set of
    stacks from block that occur earlier in the document. The intent is that the full set of
    code doesn't need to be rerun each time a block is executed because the local 'prestacks'
    are stored and used in executing the block.
    """
    def __init__(self, parent):
        self.dict   = {}
        self.SetParent(parent)
        self.ignore_keys = []

    def SetParent(self, parent):
        if parent is None:
            parent = {}
        self.parent = parent

    def __getitem__(self, name):
        if name in self.dict: return self.dict.__getitem__(name)
        if name not in self.ignore_keys:
            return self.parent[name]

    def __setitem__(self, name, value):
        if name in self.ignore_keys: self.ignore_keys.remove(self)
        self.dict.__setitem__(name,value)

    def __delitem__(self, name):
        # We don't want to pass a key to the parent if it has been deleted
        if name not in self.ignore_keys: self.ignore_keys.append(name)
        self.dict.__delitem__(name)

    def items(self):
        keys  = self.dict.keys()
        items = copy(self.dict.items())
        for key in self.parent.keys():
            if key not in keys:
                items.append((key,self.parent[key]))
        return items

    def keys(self):
        keys = copy(self.parent.keys())
        keys2 = self.dict.keys()
        for key in keys2:
            if key not in keys: keys.append(key)
        return keys

    def values(self):
        values = copy(self.parent.values())
        values2 = self.dict.values()
        for value in values2:
            if value not in values: values.append(value)
        return values

    def interitems(self):
        items = self.items()
        for item in items:
            yield item

    def iterkeys(self):
        keys = self.keys()
        for key in keys:
            yield keys

    def itervalues(self):
        values = self.values()
        for value in values:
            yield value

    def clear(self):
        self.dict.clear()

    def has_key(self, name):
        if name in self.dict:   return True
        if name in self.parent: return True
        return False

    def __str__(self):
        strng = "DictLinkWrapper:\n"
        strng += str(self.dict) + "\n"
        strng += str(self.parent)
        return strng

class StackRepository(object):
    """
    This is used to store local stacks for the different code blocks. The
    name of the code block is the key to retrieve the stack. Also, for each
    key, the name of the next block should be stored.
    """
    def __init__(self, name_handler):
        self.stacks = {}
        self.name_handler = name_handler
        name_handler.AddUpdateListener(self.__names_updated__)

    def __getitem__(self, name):
        """
        Stacks are auotmatically created if they do not already exist.
        """
        stacks = self.stacks
        if name in stacks:
            return stacks.__getitem__(name)
        else:
            parent_name = self.name_handler.GetPreviousName(name)
            if parent_name in stacks:
                parent = stacks[parent_name]
            else:
                # The parent should always be some type of dict.
                parent = {}
            wrapper = DictLinkWrapper(parent)
            self.stacks[name] = wrapper
            return wrapper

    def __setitem__(self, name, local_stack):
        self.stacks.__setitem__(name, local_stack)

    def __names_updated__(self, handler, added=[], removed=[], **kwargs):
        """
        Listens for changes to the blocks stored in the document. Updates the
        stacks based on this
        """
        for name in removed:
            if name in self.stacks: del self.stacks[name]
        self.Update(handler)

    def Update(self, handler=None):
        """
        Update the list of links for the stack linked list.
        """
        if handler is None: handler = self.name_handler
        parent = None
        names  = handler.GetNames()
        for name in names:
            if name in self.stacks: stack = self.stacks[name]
            else:
                stack = DictLinkWrapper(parent)
                self.stacks[name] = stack
            stack.SetParent(parent)
            parent = stack

    def __str__(self):
        strng = "StackRepository:\n"
        strng += str(self.stacks.keys())
        return strng

class Cursor(object):
    def __init__(self, XSCRIPTCONTEXT, block, undo_stack=None):
        self.context = XSCRIPTCONTEXT
        self.block   = block
        doc          = self.context.getDocument()
        self.cursor  = doc.Text.createTextCursor()
        self.SetCursorAfter(GetBlockTextField(self.context, block).Anchor)
        self.formula_cnt = 1
        self.undo_stack  = undo_stack
        if undo_stack is None: self.undo_stack  = []

    def __AddUndoCommand__(self, command):
        """
        For each modification to the document, an undo command should be added
        that reverses the change. The command is a function that takes no
        parameters, but stores the necessary state to reverse the changes that
        took place.
        """
        self.undo_stack.append(command)

    def ExecuteUndo(self):
        """
        Executes all the undo commands on the undo stack.
        """
        # Pull all the items off the undo stack and run them
        while len(self.undo_stack) > 0:
            try:
                command = self.undo_stack.pop()
                command()
            except Exception, e:
                # If a command fails, show the error, then continue with undo commands.
                OOTools.MsgBox(e)

    def SetCursorAfter(self, anchor):
        """
        Set the cursor after the given anchor.
        """
        if anchor is None: return False

        cursor = self.cursor
        cursor.gotoRange(anchor.getEnd(), False)
        return True

    def AddFormula(self, formula, name = None, isGlobal=True):
        """
        Adds or updates a StarMathFormula obect using the supplied formula:

        formula  - The formula in StarMath format.
        name     - A name that is used to uniquely identify the formula. This is
                   the object name found in the properties dialog for the formula.
                   If this value is not included, a local sequenced name is
                   selected. However, it is highly recommended that a name be
                   specified in order to keep formulas in sync if the script
                   block is modified.
        isGlobal - Indicates whether the formula name is global (can be accessed
                   form other blocks) or is local to this block. Local names can
                   be reused in other blocks. Global ones cannot. Local names are
                   modified to make them unique to a block. If no name is specified
                   the formula name is set to be local.
        """
        doc = self.context.getDocument()
        cursor = self.cursor
        if name is None:
            # TODO: Come up with a more robust way to keeping consistent names.
            name = "Eqn%d" % self.formula_cnt
            self.formula_cnt += 1
            isGlobal = False

        if not isGlobal:
            # Add a local prefix to keep names unique
            name = "F%s_%s" % (self.block,name)

        em = doc.getEmbeddedObjects()
        if em.hasByName(name):
            obj = em.getByName(name)
            old_formula = obj.EmbeddedObject.Formula
            def undo_update():
                "Undo update of formula content"
                obj.EmbeddedObject.Formula = old_formula
            self.__AddUndoCommand__(undo_update)

            obj.EmbeddedObject.Formula = formula
            # TODO: Figure out why it doesn't set the anchor after the formula
            self.SetCursorAfter(obj.getAnchor())
            #OOTools.inspect(anchor)
            return

        em_names = em.getElementNames()

        # The formula needs to be written as text and selected before 'InsertObjectStarMath'
        # is called, or the edit formula dialog pops up and no more inserts can take place.
        cursor.setString(formula)
        doc.CurrentController.select(cursor)
        OOTools.executeDispatch(".uno:InsertObjectStarMath")
        #sel = doc.getCurrentSelection()
        #sel.Name = name
        #OOTools.inspect(sel,name)
        OOTools.executeDispatch(".uno:Escape")
        #OOTools.inspect(cursor)
        cursor.collapseToEnd()

        em_names2 = list(em.getElementNames())
        if len(em_names2) - len(em_names) != 1:
            OOTools.MsgBox("Unable to set the Formula object name\nto %s. Formula will not update correctly." % name)

        if len(em_names) > 0:
            for dname in em_names: em_names2.remove(dname)

        obj = em.getByName(em_names2[0])
        obj.Name = name
        def undo_formula():
            "Undo insertion of StarMathObject"
            doc.Text.removeTextContent(obj)
        self.__AddUndoCommand__(undo_formula)

class CodeBlock(object):

    Local  = "Local"
    Global = "Global"

    # Order in which to return next and prev items
    DOC_ORDER  = "doc"
    EXEC_ORDER = "exec"
    TYPE_ORDER = "type"

    def __init__(self, context, block_name, stack_repo, undo_stack = []):
        """
        context    - A SCRIPTCONTEXT used to access the document
        block_name - The Name of the block.
        stack_repo - A stack repository used when executing blocks to get local variables.
        undo_stack - A list to hold undo commands (must implement append).
        """
        self.context    = context
        self.block_name = block_name
        self.undo_stack = undo_stack
        self.stack_repo = stack_repo

    def GetType(self):
        """
        Get the Block type. This is 'Global' or 'Local'.
        """
        if self.block_name.startswith('CAS_L'):
            return self.Local
        if self.block_name.startswith('CAS_G'):
            return self.Global
        return None

    def __GetCodeHeader__(self):
        """
        Get the header used in the code. This looks for a 'CAS_IMPORT' field which
        is used to define import commands. It also creates a 'doc' object which is
        an instance of the Cursor class which is available in the block. In addition,
        the current block name is accessible via 'current_block'.
        """
        block = self.block_name
        code  = ""
        doc   = self.context.getDocument()
        fm    = doc.getTextFieldMasters()
        if fm.hasByName("com.sun.star.text.FieldMaster.User.CAS_IMPORT"):
            includes = fm.getByName("com.sun.star.text.FieldMaster.User.CAS_IMPORT").Content
            code += includes + "\n\n"

        code += "#------------------------------\n"
        code += "# %s is the Active Block\n" % block
        code += "#------------------------------\n"
        code += 'current_block = "%s"\n' % block
        code += 'doc = Cursor(context, current_block, undo_stack)\n\n'
        return code

    def __GetTextFieldMaster__(self, refresh = False):
        """
        Get the TextFieldMaster associated with this CodeBlock.

        refresh - If True, then the TextFieldMaster is extracted from the
                  document. Otherwise, it may be cached.
        """
        if refresh or not hasattr(self,'TextFieldMaster'):
            doc = self.context.getDocument()
            fm = doc.getTextFieldMasters()
            name = "com.sun.star.text.FieldMaster.User." + self.block_name
            if fm.hasByName(name):
                self.TextFieldMaster = fm.getByName(name)
            else:
                return None

        return self.TextFieldMaster

    def __GetTextFields__(self, first_only=False):
        """
        Get the TextFields associated with this Code Block. In general there will
        only be one, but it is possible to have multiple in a document. If there
        are no associated TextFields (i.e. the code is not shown in the document),
        return None.

        first_only - Return only the TextField in the list.
        """
        if hasattr(self,"TextFields"):
            if first_only: return self.TextFields[0]
            return self.TextFields
        tfm = self.__GetTextFieldMaster__()
        if tfm is None: return None
        tfs = tfm.DependentTextFields
        if len(tfs) > 0:
            self.TextFields = tfs
            if first_only: return self.TextFields[0]
            return self.TextFields
        return None

    def GetBlockName(self):
        """
        Get the unique string name that identifies this block.
        """
        return self.block_name

    def GetCode(self):
        """
        Get the code associated with this code block.
        """
        tfm = self.__GetTextFieldMaster__()
        if tfm is None: return ""
        code = self.__GetCodeHeader__()
        code += tfm.Content
        return code

    def SetCode(self, code):
        """
        Set the code for this CodeBlock.

        code - A string containing the code. This is a multiline python script
               in general.
        """
        # Set the code
        tfm = self.__GetTextFieldMaster__()
        if tfm is None: return False
        old_code = tfm.Content
        tfm.Content = code
        tfs = self.__GetTextFields__()

        # Add the undo command to the undo stack
        def undo_SetCode():
            "Undo code modification for a Block"
            tfm.Content = old_code
            if tfs is not None:
                for tf in tfs: tf.update()
        self.undo_stack.append(undo_SetCode)

        if tfs is not None:
            for tf in tfs:
                tf.update()

    def Hide(self):
        """
        Hide the code section(s) for this block that are visible in the document.
        """
        tfs = self.__GetTextFields__()
        if tfs is not None:
            for tf in tfs:
                tf.IsVisible = False
                tf.update()

    def Show(self):
        """
        Show the code section(s) for this block in the document.
        """
        tfs = self.__GetTextFields__()
        if tfs is not None:
            for tf in tfs:
                tf.IsVisible = True
                tf.update()

    def GetNext(self, order="doc"):
        """
        Get the next CodeBlock object given the ordering defined by the 'order'
        variable. None if there are no CodeObjects that follow.

        order - The order in which to return items.
                'doc' indicates returning items in order they appear in document.
                'exec' indicates returning items in the order they are executed.
                'type' indicates returning items in order by type (only Local or Global).
        """

        pass

    def GetPrev(self, order="doc"):
        """
        Get the Previous CodeBlock object given the ordering defined by the 'order'
        variable. None if there are no CodeObjects that .

        order - The order in which to return items.
                'doc' indicates returning items in order they appear in document.
                'exec' indicates returning items in the order they are executed.
                'type' indicates returning items in order by type (only Local or Global).
        """
        pass

    def isBefore(self, code_block):
        """
        Compare order of two code blocks in the document.

        code_block - Another code_block object.

        Returns True if the passed in code block is before this one. Returns False if
        it is not before. Returns None if one or the other was not in the document.
        """
        self_field = self.__GetTextFields__(first_only=True)
        if self_field is None: return None
        other_field = code_block.__GetTextFields__(first_only=True)
        if other_field is None: return None

        doc = self.context.getDocument()
        if doc.Text.compareRegionStarts(other_field.Anchor, self_field.Anchor) == 1:
            return True
        else:
            return False

    def isAfter(self, code_block):
        """
        Compare order of two code blocks in the document.

        code_block - Another code_block object.

        Returns False if the passed in code block is before this one. Returns True if
        it is not before. Returns None if one or the other was not in the document.
        """
        self_field = self.__GetTextFields__(first_only=True)
        if self_field is None: return None
        other_field = code_block.__GetTextFields__(first_only=True)
        if other_field is None: return None

        doc = self.context.getDocument()
        if doc.Text.compareRegionStarts(self_field.Anchor, other_field.Anchor) == -1:
            return True
        else:
            return False

    def Remove(self):
        """
        Execute the undo_stack.
        """
        while len(undo_stack) > 0:
            try:
                command = undo_stack.pop()
                command()
            except Exception, e:
                OOTools.MsgBox(e)

        tfm = self.__GetTextFieldMaster__(refresh=True)
        if tfm is not None:
            # If there is a text field master dispose of it
            # TODO: decide if we want to delete items created by this block. Probably let the user do that.
            tfm.dispose()

    def Execute(self):
        """
        Execute the code for this block.
        """
        stack_repo = self.stack_repo
        loc = stack_repo[self.GetBlockName()]
        # Clear the stack generated last time this was run before running again
        # to avoid variable assignment conflicts.
        loc.clear()
        loc['context']    = self.context # The XSCRIPTCONTEXT object
        loc['undo_stack'] = self.undo_stack
        script = self.GetCode()
        exec script in globals(), loc

class CodeBlockFactoryImpl(object):
    def __init__(self, SCRIPTCONTEXT):
        self.SCRIPTCONTEXT = SCRIPTCONTEXT

    def GetCodeBlock(self, block_name, undo_stack = None):
        return CodeBlock(self.SCRIPTCONTEXT, block_name, StackRepositoryFactory(), undo_stack)

    def CreateCodeBlock(self, typ, code, undo_stack = None):
        """
        Creates a new code block at the current cursor location.

        SCRIPTCONTEXT - The context used to access the document.
        typ  - The type of code block to add 'Global' or 'Local'
        code - The code to add the the block. If the value is not passed in or
               is 'None' a default comment is added.
        """
        doc   = self.SCRIPTCONTEXT.getDocument()
        name  = self.GetNewBlockName(typ)
        tfm   = doc.createInstance("com.sun.star.text.FieldMaster.User")
        tfm.Name = name
        if code is None: code = "# Add Python Code Here"
        tfm.Content = code

        tf = doc.createInstance("com.sun.star.text.TextField.User")
        tf.attachTextFieldMaster(tfm)

        vc = doc.CurrentController.getViewCursor()
        doc.Text.insertTextContent(vc,tf,False)
        def undo_AddCode():
            "Remove code block and associated TextFieldMaster that were created"
            doc.Text.removeTextContent(tf)
            tfm.dispose()
        if undo_stack is not None:
            undo_stack.append(undo_AddCode)

        return self.GetCodeBlock(name, undo_stack)

    def GetNewBlockName(self, typ):
        """
        typ is 'Global' or 'Local'.
        """
        return BlockNameHandlerFactory(typ).GetNewName(self.SCRIPTCONTEXT)


BlockNameHandlerFactory = NameHandlerFactoryImpl()

_stack_repo = StackRepository(BlockNameHandlerFactory("Master"))

var_bin = {}

def CreateCodeBlockFactory(SCRIPTCONTEXT):
    return CodeBlockFactoryImpl(SCRIPTCONTEXT)

def StackRepositoryFactory():
    return _stack_repo