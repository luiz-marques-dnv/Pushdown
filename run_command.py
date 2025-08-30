from dataclasses import dataclass
import shlex

# Simple use case


def run_command(command: str) -> str:
    match command:
        case "quit":
            print("Quitting the program")
            quit()
        case "reset":
            print("Resetting the program")
        case other:
            return f"Unknown command {other!r}"

# Using split


def run_command_v2(command: str) -> str:
    cmd_parts = command.split()
    match cmd_parts:
        case ["load", filename]:
            print(f"Loading file {filename}")
        case ["save", filename]:
            print(f"Saving file {filename}")
        case ["quit" | "exit", *rest] if "--force" in rest:
            print("Sending SIGTERM to all processes and quitting")
        case ["quit" | "exit"]:
            print("Quitting the program")
            quit()
        case _:
            print(f"Unknown command {cmd_parts!r}")


@dataclass
class Command:
    name: str
    args: list[str]


def run_command_v3(command: Command):
    match command:
        case Command(name="load", args=[filename]):
            print(f"Loading file {filename}")
        case Command(name="save", args=[filename]):
            print(f"Saving file {filename}")
        case Command(name="quit" | "exit", args=["--force", *rest]):
            print("Sending SIGTERM to all processes and quitting")
            quit()
        case Command(name="quit" | "exit"):
            print("Quitting the program")
            quit()
        case _:
            print(f"Unknown command {command!r}")


def main():
    while True:
        command, *args = shlex.split(input("> "))

        run_command_v3(Command(command, args))


if __name__ == "__main__":
    main()
