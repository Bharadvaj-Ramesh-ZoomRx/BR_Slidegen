"""
SlideGen — AI-assisted PowerPoint generation and live editing.

Quick start:

    # Create a slide
    from slidegen.create import SlideBuilder
    builder = SlideBuilder()
    slide = builder.add_blank_slide()
    builder.add_clustered_bar(slide, categories, series_data, ...)
    builder.save("output.pptx")

    # Edit live (file must be open in PowerPoint)
    from slidegen.edit import LiveEditor
    with LiveEditor("output.pptx") as editor:
        editor.set_text("zrx_001", "New text")
        editor.move("zrx_002", left=5.0, top=2.0)

    # Reconcile registry before editing
    from slidegen.reconcile import reconcile
    report = reconcile("output.pptx")
"""


def __getattr__(name):
    """Lazy imports to avoid pulling heavy deps when only config is needed."""
    if name == "SlideBuilder":
        from slidegen.create import SlideBuilder
        return SlideBuilder
    if name == "LiveEditor":
        from slidegen.edit import LiveEditor
        return LiveEditor
    if name == "reconcile":
        from slidegen.reconcile import reconcile
        return reconcile
    raise AttributeError(f"module 'slidegen' has no attribute {name!r}")


__all__ = ["SlideBuilder", "LiveEditor", "reconcile"]
