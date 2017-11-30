﻿using Rubberduck.Inspections.Abstract;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Results;
using static Rubberduck.Parsing.Grammar.VBAParser;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class StepIsNotSpecifiedInspection : ParseTreeInspectionBase
    {
        public StepIsNotSpecifiedInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Warning) { }

        public override Type Type => typeof(StepIsNotSpecifiedInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.StepIsNotSpecifiedInspectionResultFormat,
                                                        result));
        }

        public override IInspectionListener Listener { get; } =
            new StepIsNotSpecifiedListener();
    }

    public class StepIsNotSpecifiedListener : VBAParserBaseListener, IInspectionListener
    {
        private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
        public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

        public QualifiedModuleName CurrentModuleName
        {
            get;
            set;
        }

        public void ClearContexts()
        {
            _contexts.Clear();
        }

        public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
        {
            StepStmtContext stepStatement = context.stepStmt();

            if (stepStatement == null)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}
