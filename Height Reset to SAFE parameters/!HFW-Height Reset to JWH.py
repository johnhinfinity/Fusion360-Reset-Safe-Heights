import adsk.core, adsk.cam, traceback

def _get_cam(app: adsk.core.Application):
    """Return the CAM product for the active document."""
    doc = app.activeDocument
    if not doc:
        return None

    products = doc.products
    cam = products.itemByProductType('CAMProductType')
    if cam:
        return adsk.cam.CAM.cast(cam)

    # Fallback: if Manufacture is currently active
    return adsk.cam.CAM.cast(app.activeProduct)


def run(context):
    app = adsk.core.Application.get()
    ui = app.userInterface

    try:
        cam = _get_cam(app)
        if not cam:
            ui.messageBox('No CAM product found.\nOpen a document with CAM and switch to Manufacture.')
            return

        ops = cam.allOperations
        if ops.count == 0:
            ui.messageBox('No operations found in this document.')
            return

        changed_ops = 0
        changed_params = 0

        for op in ops:
            params = op.parameters
            touched = False

            # Look up each parameter by its exact name from your dump
            clr_mode   = params.itemByName('clearanceHeight_mode')
            clr_offset = params.itemByName('clearanceHeight_offset')

            rtr_mode   = params.itemByName('retractHeight_mode')
            rtr_offset = params.itemByName('retractHeight_offset')

            feed_mode   = params.itemByName('feedHeight_mode')
            feed_offset = params.itemByName('feedHeight_offset')

            try:
                # Clearance: From Retract height, 0.5 in
                if clr_mode:
                    clr_mode.expression = "'from retract height'"
                    changed_params += 1
                    touched = True
                if clr_offset:
                    clr_offset.expression = "0.5in"
                    changed_params += 1
                    touched = True

                # Retract: From Feed height, 2 in
                if rtr_mode:
                    rtr_mode.expression = "'from feed height'"
                    changed_params += 1
                    touched = True
                if rtr_offset:
                    rtr_offset.expression = "2in"
                    changed_params += 1
                    touched = True

                # Feed: From Top height, 0.2 in  (Fusion uses 'from top')
                if feed_mode:
                    feed_mode.expression = "'from top'"
                    changed_params += 1
                    touched = True
                if feed_offset:
                    feed_offset.expression = "0.2in"
                    changed_params += 1
                    touched = True

            except:
                # If a weird operation type rejects something, just skip it
                pass

            if touched:
                changed_ops += 1

        ui.messageBox(
            f'Finished.\n'
            f'Modified {changed_params} parameters on {changed_ops} operations out of {ops.count}.'
        )

    except:
        ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
