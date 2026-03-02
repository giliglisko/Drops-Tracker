using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace Drops_Tracker.Helpers
{
    public class DragAdorner : Adorner
    {
        private readonly UIElement _child;
        private Point _offset;

        public DragAdorner(UIElement adornedElement, UIElement child, Point offset) : base(adornedElement)
        {
            _child = child;
            _offset = offset;
            IsHitTestVisible = false;
            AddVisualChild(_child);
        }

        public void UpdatePosition(Point position)
        {
            _offset = position;
            InvalidateArrange();
        }

        protected override int VisualChildrenCount => 1;

        protected override Visual GetVisualChild(int index) => _child;

        protected override Size MeasureOverride(Size constraint)
        {
            _child.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            return _child.DesiredSize;
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            _child.Arrange(new Rect(_child.DesiredSize));
            return finalSize;
        }

        public override GeneralTransform GetDesiredTransform(GeneralTransform transform)
        {
            var result = new GeneralTransformGroup();
            result.Children.Add(new TranslateTransform(_offset.X - 30, _offset.Y - 30));
            var baseTransform = base.GetDesiredTransform(transform);
            if (baseTransform != null)
            {
                result.Children.Add(baseTransform);
            }
            return result;
        }
    }
}
