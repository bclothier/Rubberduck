﻿using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Rename;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using MockFactory = RubberduckTests.Mocks.MockFactory;

namespace RubberduckTests.Refactoring
{
    [TestClass]
    public class RenameTests : VbeTestBase
    {
        [TestMethod]
        public void RenameRefactoring_RenameSub()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
End Sub";
            var selection = new Selection(2, 12, 2, 12); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
End Sub";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameSub_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub

Private Sub Goo()
    Foo
End Sub
";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Hoo()
End Sub

Private Sub Goo()
    Hoo
End Sub
";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameVariable_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
    Dim val1 As Integer
    val1 = val1 + 5
End Sub";
            var selection = new Selection(2, 12, 2, 12); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo()
    Dim val2 As Integer
    val2 = val2 + 5
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "val2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameParameter_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Sub Foo(ByVal arg1 As String)
    arg1 = ""test""
End Sub";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Foo(ByVal arg2 As String)
    arg2 = ""test""
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "arg2" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndSetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo(ByVal arg1 As Integer) 
End Property

Private Property Set Foo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Goo(ByVal arg1 As Integer) 
End Property

Private Property Set Goo(ByVal arg1 As Integer, ByVal arg2 As String) 
End Property";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameGetterAndLetter()
        {
            //Input
            const string inputCode =
@"Private Property Get Foo() 
End Property

Private Property Let Foo(ByVal arg1 As String) 
End Property";
            var selection = new Selection(1, 25, 1, 25); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Property Get Goo() 
End Property

Private Property Let Goo(ByVal arg1 As String) 
End Property";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function";
            var selection = new Selection(1, 21, 1, 21); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Goo() As Boolean
    Goo = True
End Function";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameFunction_UpdatesReferences()
        {
            //Input
            const string inputCode =
@"Private Function Foo() As Boolean
    Foo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Foo()
End Sub
";
            var selection = new Selection(1, 21, 1, 21); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Function Hoo() As Boolean
    Hoo = True
End Function

Private Sub Goo()
    Dim var1 As Boolean
    var1 = Hoo()
End Sub
";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Hoo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RefactorWithDeclaration()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";
            var selection = new Selection(1, 15, 1, 15); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode =
@"Private Sub Goo()
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(model.Target);

            //Assert
            Assert.AreEqual(expectedCode, module.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameInterface()
        {
            //Input
            const string inputCode1 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(1, 22, 1, 22); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";
            const string expectedCode2 =
@"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var vbe = MockFactory.CreateVbeMock();
            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_RenameEvent()
        {
            //Input
            const string inputCode1 =
@"Public Event Foo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string inputCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Foo(ByVal i As Integer, ByVal s As String)
End Sub";

            var selection = new Selection(1, 16, 1, 16); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Public Event Goo(ByVal arg1 As Integer, ByVal arg2 As String)";
            const string expectedCode2 =
@"Private WithEvents abc As Class1

Private Sub abc_Goo(ByVal i As Integer, ByVal s As String)
End Sub";

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "Class2",
                vbext_ComponentType.vbext_ct_ClassModule);

            var vbe = MockFactory.CreateVbeMock();
            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, null) { NewName = "Goo" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(qualifiedSelection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_InterfaceRenamed_AcceptPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 27, 3, 27); //startLine, startCol, endLine, endCol

            //Expectation
            const string expectedCode1 =
@"Implements IClass1

Private Sub IClass1_DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";
            const string expectedCode2 =
@"Public Sub DoNothing(ByVal a As Integer, ByVal b As String)
End Sub";

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var vbe = MockFactory.CreateVbeMock();
            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var module1 = project.Object.VBComponents.Item(0).CodeModule;
            var module2 = project.Object.VBComponents.Item(1).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.Yes);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, messageBox.Object) { NewName = "DoNothing" };

            //SetupFactory
            var factory = SetupFactory(model);

            //Act
            var refactoring = new RenameRefactoring(factory.Object);
            refactoring.Refactor(model.Selection);

            //Assert
            Assert.AreEqual(expectedCode1, module1.Lines());
            Assert.AreEqual(expectedCode2, module2.Lines());
        }

        [TestMethod]
        public void RenameRefactoring_InterfaceRenamed_RejectPrompt()
        {
            //Input
            const string inputCode1 =
@"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";
            const string inputCode2 =
@"Public Sub DoSomething(ByVal a As Integer, ByVal b As String)
End Sub";

            var selection = new Selection(3, 23, 3, 27); //startLine, startCol, endLine, endCol

            //Arrange
            var component1 = CreateMockComponent(inputCode1, "Class1",
                vbext_ComponentType.vbext_ct_ClassModule);
            var component2 = CreateMockComponent(inputCode2, "IClass1",
                vbext_ComponentType.vbext_ct_ClassModule);

            var vbe = MockFactory.CreateVbeMock();
            var project = CreateMockProject("VBEProject", vbext_ProjectProtection.vbext_pp_none,
                new List<Mock<VBComponent>>() { component1, component2 });
            var parseResult = new RubberduckParser().Parse(project.Object);

            var qualifiedSelection = GetQualifiedSelection(selection);

            var messageBox = new Mock<IMessageBox>();
            messageBox.Setup(
                m => m.Show(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<MessageBoxButtons>(), It.IsAny<MessageBoxIcon>()))
                .Returns(DialogResult.No);

            var model = new RenameModel(vbe.Object, parseResult, qualifiedSelection, messageBox.Object);
            Assert.AreEqual(null, model.Target);
        }
        
        [TestMethod]
        public void Rename_PresenterIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var module = project.Object.VBComponents.Item(0).CodeModule;
            var parseResult = new RubberduckParser().Parse(project.Object);

            var factory = new RenamePresenterFactory(vbe.Object, null, parseResult, null);

            //act
            var refactoring = new RenameRefactoring(factory);
            refactoring.Refactor();

            Assert.AreEqual(inputCode, module.Lines());
        }

        [TestMethod]
        public void Presenter_TargetIsNull()
        {
            //Input
            const string inputCode =
@"
Private Sub Foo(ByVal arg1 As Integer, ByVal arg2 As String)
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);
            
            var factory = new RenamePresenterFactory(vbe.Object, null,
                parseResult, null);

            var presenter = factory.Create();

            Assert.AreEqual(null, presenter.Show());
        }
        
        [TestMethod]
        public void Factory_SelectionIsNull()
        {
            //Input
            const string inputCode =
@"Private Sub Foo()
End Sub";

            //Arrange
            var vbe = MockFactory.CreateVbeMock();
            var project = SetupMockProject(inputCode);
            var parseResult = new RubberduckParser().Parse(project.Object);

            var editor = new Mock<IActiveCodePaneEditor>();
            editor.Setup(e => e.GetSelection()).Returns((QualifiedSelection?)null);

            var factory = new RenamePresenterFactory(vbe.Object, null, parseResult, null);

            var presenter = factory.Create();
            Assert.AreEqual(null, presenter.Show());
        }

        #region setup
        private static Mock<IRefactoringPresenterFactory<IRenamePresenter>> SetupFactory(RenameModel model)
        {
            var presenter = new Mock<IRenamePresenter>();
            presenter.Setup(p => p.Show()).Returns(model);
            presenter.Setup(p => p.Show(It.IsAny<Declaration>())).Returns(model);

            var factory = new Mock<IRefactoringPresenterFactory<IRenamePresenter>>();
            factory.Setup(f => f.Create()).Returns(presenter.Object);
            return factory;
        }

        #endregion
    }
}