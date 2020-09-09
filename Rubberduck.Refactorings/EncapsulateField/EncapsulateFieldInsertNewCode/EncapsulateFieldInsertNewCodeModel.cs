﻿using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldInsertNewCode
{
    public class EncapsulateFieldInsertNewCodeModel : IRefactoringModel
    {
        public EncapsulateFieldInsertNewCodeModel(IEnumerable<IEncapsulateFieldCandidate> selectedFieldCandidates)
        {
            SelectedFieldCandidates = selectedFieldCandidates.ToList();
            if (SelectedFieldCandidates.Any())
            {
                QualifiedModuleName = SelectedFieldCandidates.First().QualifiedModuleName;
            }
        }

        public INewContentAggregator NewContentAggregator { set; get; }

        public QualifiedModuleName QualifiedModuleName { get; } = new QualifiedModuleName();

        public IReadOnlyCollection<IEncapsulateFieldCandidate> SelectedFieldCandidates { get; }
    }
}
