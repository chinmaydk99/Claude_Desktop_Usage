import asyncio
import base64
import os
from dataclasses import dataclass, fields, replace
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal, Optional, TypedDict, List, Dict, cast
from datetime import datetime
import anthropic
from anthropic.types import MessageParam
import pyautogui
from io import BytesIO
import win32gui
import win32con
import keyboard
from PIL import Image
import platform

print("""NOTE: Please be very careful running this script!! this runs locally on your machine(NO SANDBOX) There is an artificial delay in script before each action so you can review them(default 5 seconds). KEEP A WATCHFUL EYE ON IT! AND STOP THE SCRIPT DURING WAIT TIME IF IT TRIES TO DO SOMETHING YOU DONT WANT. BY RUNNING THIS SCRIPT YOU ASSUME THE RESPONSIBILITY OF THE OUTCOMES""")

# This is to allow the user to confirm or abort actions with some time to think
WAIT_BEFORE_ACTION: Optional[float] = None  # Set to None to disable waiting, or to a number of seconds

# Configure PyAutoGUI
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

# Actions defined in the official repository
Action = Literal[
    "key",
    "type",
    "mouse_move",
    "left_click",
    "left_click_drag",
    "right_click",
    "middle_click",
    "double_click",
    "screenshot",
    "cursor_position",
]

Command = Literal[
    "view",
    "create",
    "str_replace",
    "insert",
    "undo_edit",
]

class Resolution(TypedDict):
    width: int
    height: int

MAX_SCALING_TARGETS: dict[str, Resolution] = {
    "XGA": Resolution(width=1024, height=768),  # 4:3
    "WXGA": Resolution(width=1280, height=800),  # 16:10
    "FWXGA": Resolution(width=1366, height=768),  # ~16:9
}

@dataclass(frozen=True)
class ToolResult:
    """Base result type for all tools"""
    output: Optional[str] = None
    error: Optional[str] = None
    base64_image: Optional[str] = None
    system: Optional[str] = None

    def __bool__(self):
        return any(getattr(self, field.name) for field in fields(self))
    
    def replace(self, **kwargs):
        return replace(self, **kwargs)

class ToolError(Exception):
    def __init__(self, message: str):
        self.message = message
        super().__init__(message)

class ComputerTool:
    """Windows-compatible computer interaction tool"""
    name = "computer"
    api_type = "computer_20241022"
    _screenshot_delay = 1.0
    
    def __init__(self):
        self.screen_width, self.screen_height = pyautogui.size()
        target_res = MAX_SCALING_TARGETS["XGA"]
        self.width = target_res["width"]
        self.height = target_res["height"]
        
    def to_params(self):
        return {
            "type": self.api_type,
            "name": self.name,
            "display_width_px": self.width,
            "display_height_px": self.height,
            "display_number": 1,
        }

    async def __call__(self, action: Action, text: Optional[str] = None,
                       coordinate: Optional[tuple[int, int] | List[int]] = None, **kwargs):
        try:
            # Get action description
            action_desc = self._get_action_description(action, text, coordinate)
            print(f"\nPending Action: {action_desc}")
            
            # Wait if WAIT_BEFORE_ACTION is set
            if WAIT_BEFORE_ACTION is not None:
                print(f"Waiting {WAIT_BEFORE_ACTION} seconds before executing action. Press Ctrl+C to abort...")
                await asyncio.sleep(WAIT_BEFORE_ACTION)

            # Convert list coordinates to tuple if necessary
            if isinstance(coordinate, list) and len(coordinate) == 2:
                coordinate = tuple(coordinate)

            # Scale coordinates if provided
            if coordinate:
                if not isinstance(coordinate, tuple) or len(coordinate) != 2:
                    raise ToolError(f"Invalid coordinate format: {coordinate}")
                try:
                    x, y = int(coordinate[0]), int(coordinate[1])
                    x, y = self._scale_coordinates(x, y)
                    coordinate = (x, y)
                except (ValueError, TypeError):
                    raise ToolError(f"Invalid coordinate values: {coordinate}")

            # Execute action
            if action in ("mouse_move", "left_click_drag"):
                if not coordinate:
                    raise ToolError(f"coordinate required for {action}")
                x, y = coordinate
                print(f"Moving to scaled coordinates: ({x}, {y})")
                if action == "mouse_move":
                    pyautogui.moveTo(x, y)
                else:
                    pyautogui.dragTo(x, y, button='left')
                
            elif action in ("key", "type"):
                if not text:
                    raise ToolError(f"text required for {action}")
                print(f"Sending text: {text}")
                if action == "key":
                    keyboard.send(text)
                else:
                    pyautogui.write(text, interval=0.01)
                
            elif action in ("left_click", "right_click", "middle_click", "double_click"):
                print(f"Performing {action}")
                click_map = {
                    "left_click": lambda: pyautogui.click(button='left'),
                    "right_click": lambda: pyautogui.click(button='right'),
                    "middle_click": lambda: pyautogui.click(button='middle'),
                    "double_click": lambda: pyautogui.doubleClick()
                }
                click_map[action]()
                
            elif action == "cursor_position":
                x, y = pyautogui.position()
                scaled_x, scaled_y = self._inverse_scale_coordinates(x, y)
                return ToolResult(output=f"X={scaled_x},Y={scaled_y}")
            
            elif action == "screenshot":
                return await self._take_screenshot()

            # Always take a screenshot after any action (except cursor_position)
            if action != "cursor_position":
                await asyncio.sleep(self._screenshot_delay)
                result = await self._take_screenshot()
                if result.error:
                    raise ToolError(result.error)
                return result
            
        except Exception as e:
            error_msg = f"Action failed: {str(e)}"
            print(f"\nError: {error_msg}")
            return ToolResult(error=error_msg)

    def _scale_coordinates(self, x: int, y: int) -> tuple[int, int]:
        """Scale coordinates from XGA to actual screen resolution"""
        scaled_x = int(x * (self.screen_width / self.width))
        scaled_y = int(y * (self.screen_height / self.height))
        return scaled_x, scaled_y

    def _inverse_scale_coordinates(self, x: int, y: int) -> tuple[int, int]:
        """Scale coordinates from actual screen resolution to XGA"""
        scaled_x = int(x * (self.width / self.screen_width))
        scaled_y = int(y * (self.height / self.screen_height))
        return scaled_x, scaled_y

    async def _take_screenshot(self) -> ToolResult:
        try:
            screenshot = pyautogui.screenshot()
            if screenshot.size != (self.width, self.height):
                screenshot = screenshot.resize(
                    (self.width, self.height), 
                    Image.Resampling.LANCZOS
                )
            
            buffered = BytesIO()
            screenshot.save(buffered, format="PNG", optimize=True)
            img_str = base64.b64encode(buffered.getvalue()).decode()
            
            return ToolResult(base64_image=img_str)
        except Exception as e:
            return ToolResult(error=f"Screenshot failed: {str(e)}")

    def _get_action_description(self, action: Action, text: Optional[str],
                              coordinate: Optional[tuple[int, int] | List[int]]) -> str:
        if action in ("mouse_move", "left_click_drag"):
            return f"{action.replace('_', ' ').title()} to coordinates: {coordinate}"
        elif action in ("key", "type"):
            return f"{action.title()} text: '{text}'"
        elif action in ("left_click", "right_click", "middle_click", "double_click"):
            return f"Perform {action.replace('_', ' ')}"
        elif action == "screenshot":
            return "Take a screenshot"
        elif action == "cursor_position":
            return "Get current cursor position"
        return f"Unknown action: {action}"

class EditTool:
    """File editing tool"""
    name = "str_replace_editor"
    api_type = "text_editor_20241022"
    
    def __init__(self):
        self._file_history = {}

    def to_params(self):
        return {
            "type": self.api_type,
            "name": self.name
        }

    async def __call__(self, command: Command, path: str, file_text: Optional[str] = None,
                       old_str: Optional[str] = None, new_str: Optional[str] = None,
                       insert_line: Optional[int] = None, view_range: Optional[List[int]] = None,
                       **kwargs):
        try:
            path_obj = Path(path)
            
            if not path_obj.is_absolute():
                suggested_path = Path.cwd() / path
                raise ToolError(f"Path must be absolute. Did you mean: {suggested_path}?")

            if command == "view":
                content = path_obj.read_text(encoding='utf-8')
                if view_range:
                    lines = content.splitlines()
                    start, end = view_range
                    content = '\n'.join(lines[start-1:end])
                return ToolResult(output=self._format_output(content, path))
            
            elif command == "create":
                if not file_text:
                    raise ToolError("file_text required for create")
                if path_obj.exists():
                    raise ToolError(f"File already exists: {path}")
                path_obj.write_text(file_text, encoding='utf-8')
                self._file_history[path] = [file_text]
                return ToolResult(output=f"File created at {path}")

            elif command == "str_replace":
                if not old_str:
                    raise ToolError("old_str required for str_replace")
                content = path_obj.read_text(encoding='utf-8')
                occurrences = content.count(old_str)
                if occurrences == 0:
                    raise ToolError(f"String '{old_str}' not found in file")
                if occurrences > 1:
                    raise ToolError(f"Multiple occurrences ({occurrences}) of '{old_str}' found")
                new_content = content.replace(old_str, new_str or "")
                self._file_history[path] = self._file_history.get(path, []) + [content]
                path_obj.write_text(new_content, encoding='utf-8')
                return ToolResult(output=self._format_output(new_content, path))

            elif command == "insert":
                if insert_line is None or not new_str:
                    raise ToolError("insert_line and new_str required for insert")
                content = path_obj.read_text(encoding='utf-8')
                lines = content.splitlines()
                if not (0 <= insert_line <= len(lines)):
                    raise ToolError(f"Invalid line number: {insert_line}")
                lines.insert(insert_line, new_str)
                new_content = '\n'.join(lines)
                self._file_history[path] = self._file_history.get(path, []) + [content]
                path_obj.write_text(new_content, encoding='utf-8')
                return ToolResult(output=self._format_output(new_content, path))

            raise ToolError(f"Invalid command: {command}")
            
        except Exception as e:
            if isinstance(e, ToolError):
                raise
            raise ToolError(f"File operation failed: {str(e)}")

    def _format_output(self, content: str, path: str) -> str:
        numbered_lines = [f"{i+1:6}\t{line}" for i, line in enumerate(content.splitlines())]
        return f"Content of {path}:\n" + '\n'.join(numbered_lines)

class ComputerControlAPI:
    def __init__(self, api_key: str, model: str = "claude-3-5-sonnet-20241022"):
        self.client = anthropic.Anthropic(api_key=api_key)
        self.model = model
        self.computer_tool = ComputerTool()
        self.edit_tool = EditTool()
        
    async def run_conversation(self):
        messages = []
        system = self._get_system_prompt()
        print("\nComputer Control Assistant Initialized")
        print("Display configured for XGA resolution (1024x768)")
        
        try:
            while True:
                if not messages:
                    user_input = input("\nWhat would you like me to do? (type 'exit' to quit): ").strip()
                    if user_input.lower() == 'exit':
                        break
                    messages.append({
                        "role": "user",
                        "content": [{"type": "text", "text": user_input}]
                    })

                print("\nProcessing request...")
                response = self.client.beta.messages.create(
                    model=self.model,
                    messages=cast(List[MessageParam], messages),
                    system=system,
                    tools=[self.computer_tool.to_params(), self.edit_tool.to_params()],
                    max_tokens=4096,
                    extra_headers={"anthropic-beta": "computer-use-2024-10-22"}
                )

                # Process assistant's response
                tool_calls = []
                print("\nAssistant's Plan:")
                
                # Handle response content properly
                for block in response.content:
                    # Print text responses
                    if hasattr(block, 'text'):
                        print(f"\n{block.text}")
                    # Collect tool calls
                    if hasattr(block, 'type') and block.type == 'tool_use':
                        tool_calls.append({
                            "name": block.name,
                            "input": block.input,
                            "id": block.id
                        })
                        print(f"\n[Planning: {block.name} - {block.input}]")

                messages.append({
                    "role": "assistant",
                    "content": response.content
                })

                if not tool_calls:
                    user_continue = input("\nNo actions to perform. Continue? (yes/no): ").lower()
                    if user_continue != 'yes':
                        break
                    messages = []
                    continue

                # Execute tools and collect results
                tool_results = []
                for tool_call in tool_calls:
                    try:
                        tool_name = tool_call.get("name")
                        tool_input = tool_call.get("input", {})
                        
                        # Select appropriate tool
                        tool = self.computer_tool if tool_name == "computer" else self.edit_tool
                        
                        print(f"\nExecuting {tool_name} with input: {tool_input}")
                        
                        # Execute tool
                        result = await tool(**tool_input)
                        
                        # Format and collect results
                        formatted_result = {
                            "type": "tool_result",
                            "tool_use_id": tool_call.get("id"),
                            "is_error": bool(result.error),
                            "content": self._format_tool_result(result)
                        }
                        
                        tool_results.append(formatted_result)
                        
                        # Show results to user
                        if result.error:
                            print(f"\nError: {result.error}")
                        elif result.output:
                            print(f"\nResult: {result.output}")
                        
                        # Small delay between actions
                        await asyncio.sleep(0.5)
                        
                    except Exception as e:
                        print(f"\nError executing tool: {str(e)}")
                        tool_results.append({
                            "type": "tool_result",
                            "tool_use_id": tool_call.get("id"),
                            "is_error": True,
                            "content": [{"type": "text", "text": f"Tool execution failed: {str(e)}"}]
                        })

                if tool_results:
                    messages.append({
                        "role": "user",
                        "content": tool_results
                    })

        except KeyboardInterrupt:
            print("\nOperation cancelled by user")
        except Exception as e:
            print(f"\nAn error occurred: {str(e)}")
            if DEBUG:
                import traceback
                traceback.print_exc()
        finally:
            print("\nThank you for using Computer Control Assistant!")

    def _format_tool_result(self, result: ToolResult) -> List[Dict[str, Any]]:
        content = []
        
        if result.error:
            return [{"type": "text", "text": result.error}]
        
        if result.system:
            content.append({"type": "text", "text": f"<system>{result.system}</system>"})
            
        if result.output:
            content.append({"type": "text", "text": result.output})
            
        if result.base64_image:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": result.base64_image
                }
            })
            
        return content

    def _get_system_prompt(self) -> str:
        return f"""<SYSTEM_CAPABILITY>
        * You are utilizing a Windows {platform.machine()} machine.
        * You can control the computer through mouse movements, clicks, and keyboard input.
        * The display is configured for XGA resolution (1024x768) for consistency.
        * All coordinates you receive and send should be in XGA resolution - they will be automatically scaled.
        * After each action you'll receive a screenshot to confirm the result.
        * Each action requires user confirmation via Enter key before execution.
        * You can use both keyboard.send() for special keys and type() for text input.
        </SYSTEM_CAPABILITY>

        <IMPORTANT>
        * Always take small, deliberate steps and verify the results through screenshots.
        * When using Firefox, if a startup wizard appears, ignore it and click directly on the address bar.
        * For text-heavy pages, consider downloading the content and using the text editor tool.
        * Before each action, clearly explain what you're about to do and why.
        * If an action fails, try alternative approaches or ask for user guidance.
        * Keep track of window focus and mouse position for accurate interactions.
        * Use 'key' action for special keys like 'ctrl', 'alt', 'tab', 'enter', etc.
        * Use 'type' action for regular text input.
        </IMPORTANT>

        <SPECIAL_KEYS>
        * Navigation: 'left', 'right', 'up', 'down', 'home', 'end', 'pageup', 'pagedown'
        * Editing: 'backspace', 'delete', 'tab', 'enter'
        * Modifiers: 'ctrl', 'alt', 'shift'
        * Function: 'f1' through 'f12'
        * Combinations: Use '+' to combine keys, e.g., 'ctrl+c', 'alt+tab'
        </SPECIAL_KEYS>

        The current date is {datetime.now().strftime('%A, %B %d, %Y')}.
        """

async def main():
    """Main entry point for the Computer Control application"""
    print("\nWindows Computer Control Assistant")
    print("==================================")
    
    try:
        # Get API key from environment or user input
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            api_key = input("Please enter your Anthropic API key: ").strip()
            if not api_key:
                print("API key is required.")
                return
            os.environ["ANTHROPIC_API_KEY"] = api_key

        # Get wait time from user
        global WAIT_BEFORE_ACTION
        wait_input = input("Enter wait time before actions (in seconds), or press Enter for default (5 seconds), or type 'none' for no wait: ").strip().lower()
        if wait_input == 'none':
            WAIT_BEFORE_ACTION = None
        elif wait_input:
            try:
                WAIT_BEFORE_ACTION = float(wait_input)
            except ValueError:
                print("Invalid input. Using default wait time of 5 seconds.")
        print(f"Wait time set to: {WAIT_BEFORE_ACTION if WAIT_BEFORE_ACTION is not None else 'No wait'}")

        # Initialize and run the API
        api = ComputerControlAPI(api_key=api_key)
        await api.run_conversation()
        
    except KeyboardInterrupt:
        print("\nProgram terminated by user.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {str(e)}")
        if DEBUG:
            import traceback
            traceback.print_exc()
    finally:
        print("\nProgram ended.")

# Global settings
DEBUG = False  # Set to True for detailed error messages



if __name__ == "__main__":
    # Set up basic configuration
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.1  # 100ms pause between actions
    
    # Run the main async loop
    asyncio.run(main())

