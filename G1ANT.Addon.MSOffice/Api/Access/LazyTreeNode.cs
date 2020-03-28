using G1ANT.Addon.MSOffice.Models.Access;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace G1ANT.Addon.MSOffice.Api.Access
{
    internal class LazyTreeNode : TreeNode
    {
        private readonly List<Func<IEnumerable<TreeNode>>> treeNodeFactories;

        public LazyTreeNode(string text) : base(text)
        {
            this.treeNodeFactories = new List<Func<IEnumerable<TreeNode>>>();

            Nodes.Add("");
        }

        public LazyTreeNode(string text, object model) : this(text)
        {
            this.Tag = model;
            if (model is IDetailedNameModel detailedModel)
                this.ToolTipText = detailedModel.ToDetailedString();
        }

        public LazyTreeNode(object model) : this(model.ToString(), model)
        { }


        public LazyTreeNode Add(TreeNode treeNode)
        {
            treeNodeFactories.Add(() => new List<TreeNode>() { treeNode });
            return this;
        }

        public LazyTreeNode AddRange(params Func<IEnumerable<TreeNode>>[] treeNodeFactories)
        {
            this.treeNodeFactories.AddRange(treeNodeFactories);
            return this;
        }

        internal LazyTreeNode Add(Func<TreeNode> treeNode)
        {
            treeNodeFactories.Add(() => new List<TreeNode>() { treeNode() });
            return this;
        }

        private bool IsEmpty()
        {
            return Nodes.Count == 1 && Nodes[0].Text == "";
        }

        public void LoadLazyChildren()
        {
            if (IsEmpty())
            {
                Nodes.Clear();
                treeNodeFactories.ToList().ForEach(f => Nodes.AddRange(f().ToArray()));
            }
        }

        internal LazyTreeNode EmptyChildren()
        {
            Nodes.Clear();
            return this;
        }
    }
}
