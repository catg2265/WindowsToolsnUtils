function Start-BatteryTest{
    param(
        [object]$Progress
    )

    # Add this where the logic of progress bar is
    & $Progress.Update.Invoke(
        "Battery Test",
        "Discharging...",
        $i
    )

    # Run this to complete the progress bar
    & $Progress.Complete.Invoke()
}