"""Discover every workflow class in the `worker.workflows` package and serve them."""

import asyncio
import signal

import mistralai.workflows as workflows


async def main() -> None:
    workflow_classes = workflows.discover_all_workflows_in_package("worker.workflows")
    names = [workflows.get_workflow_definition(c).name for c in workflow_classes]
    print(f"workflow discovery complete: {names}", flush=True)

    served = asyncio.current_task()
    stopped_by: signal.Signals | None = None

    def _stop(stop_signal: signal.Signals) -> None:
        nonlocal stopped_by
        if stopped_by is not None:
            return
        stopped_by = stop_signal
        if served is not None:
            served.cancel()

    loop = asyncio.get_running_loop()
    for stop_signal in (signal.SIGTERM, signal.SIGINT):
        loop.add_signal_handler(stop_signal, _stop, stop_signal)
    try:
        await workflows.run_worker(workflow_classes)
    except asyncio.CancelledError:
        if stopped_by is None:
            raise


if __name__ == "__main__":
    asyncio.run(main())
