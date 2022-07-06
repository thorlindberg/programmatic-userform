This Excel project consists of a class and supporting macros developed with Visual Basic for Applications. The purpose of this project is to provide a framework for generating UserForms programmatically and dynamically, meaning they can be instantiated, initialised, displayed and destroyed, with their entire lifecycle existing in the code.

This implementation provides an alternative to instantiating and initialising UserForms through files, which is difficult to develop and document because they cannot be previewed with or commented in the code. This situation creates barriers for collaboration between multiple developers working concurrently on the same Excel project.

<br>

**Life cycle**

Excel provides support for multiple co-existing lifecycles of its UserForms. These lifecycles include static forms as files, dynamic forms as files with dynamic changes during runtime, and designer forms programmatically generated during runtime. As these lifecycles can co-exist, a UserForm can be multiple, all or one of these types.

<br>

**Use case**

A programmatic UserForm is instantiated as an object in a macro or class, and can then be displayed as needed. The object is a Form of type `UserInterface` which provides a series of methods to dynamically initialise the contents of the form. As a UserForm is generated through these methods, the class automatically adjusts the size and layout to fit the content added.

<br>

**Implementation**

The class and macros can be imported into a new Excel document or along with the existing macros of an established project. This should not have any impact on existing macros, and can be utilised extensibly to generate and display UserForms. These forms are programmatically generated, and as such they have no source file, which means a form is destroyed once it is closed through user interaction.
