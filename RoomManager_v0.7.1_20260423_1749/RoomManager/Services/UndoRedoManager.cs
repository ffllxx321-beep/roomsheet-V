using RoomManager.Models;
using System.Collections.Generic;
using System.Linq;

namespace RoomManager.Services;

/// <summary>
/// 撤销/重做管理器
/// </summary>
public class UndoRedoManager
{
    private readonly Stack<RoomOperation> _undoStack = new();
    private readonly Stack<RoomOperation> _redoStack = new();
    private readonly int _maxHistorySize;

    /// <summary>
    /// 是否可以撤销
    /// </summary>
    public bool CanUndo => _undoStack.Count > 0;

    /// <summary>
    /// 是否可以重做
    /// </summary>
    public bool CanRedo => _redoStack.Count > 0;

    /// <summary>
    /// 撤销栈大小
    /// </summary>
    public int UndoCount => _undoStack.Count;

    /// <summary>
    /// 重做栈大小
    /// </summary>
    public int RedoCount => _redoStack.Count;

    /// <summary>
    /// 操作执行事件
    /// </summary>
    public event EventHandler<OperationEventArgs>? OperationExecuted;

    public UndoRedoManager(int maxHistorySize = 50)
    {
        _maxHistorySize = maxHistorySize;
    }

    /// <summary>
    /// 执行操作并记录到撤销栈
    /// </summary>
    public void ExecuteOperation(RoomOperation operation)
    {
        // 执行操作
        operation.Execute();

        // 添加到撤销栈
        _undoStack.Push(operation);

        // 清空重做栈
        _redoStack.Clear();

        // 限制历史大小（移除最旧的操作）
        if (_undoStack.Count > _maxHistorySize)
        {
            // 将栈转为数组，移除最后一个（最旧的），再重新入栈
            var items = _undoStack.ToArray();
            _undoStack.Clear();
            // 跳过最后一个元素（最旧的），从新到旧重新入栈
            for (int i = 0; i < items.Length - 1; i++)
            {
                _undoStack.Push(items[i]);
            }
        }

        OperationExecuted?.Invoke(this, new OperationEventArgs
        {
            Operation = operation,
            Type = OperationType.Execute
        });
    }

    /// <summary>
    /// 撤销最近操作
    /// </summary>
    public RoomOperation? Undo()
    {
        if (!CanUndo) return null;

        var operation = _undoStack.Pop();
        operation.Undo();
        _redoStack.Push(operation);

        OperationExecuted?.Invoke(this, new OperationEventArgs
        {
            Operation = operation,
            Type = OperationType.Undo
        });

        return operation;
    }

    /// <summary>
    /// 重做最近撤销的操作
    /// </summary>
    public RoomOperation? Redo()
    {
        if (!CanRedo) return null;

        var operation = _redoStack.Pop();
        operation.Execute();
        _undoStack.Push(operation);

        OperationExecuted?.Invoke(this, new OperationEventArgs
        {
            Operation = operation,
            Type = OperationType.Redo
        });

        return operation;
    }

    /// <summary>
    /// 清空所有历史
    /// </summary>
    public void Clear()
    {
        _undoStack.Clear();
        _redoStack.Clear();
    }

    /// <summary>
    /// 获取撤销历史描述
    /// </summary>
    public List<string> GetUndoHistory()
    {
        return _undoStack.Select(op => op.Description).ToList();
    }

    /// <summary>
    /// 获取重做历史描述
    /// </summary>
    public List<string> GetRedoHistory()
    {
        return _redoStack.Select(op => op.Description).ToList();
    }
}

/// <summary>
/// 房间操作基类
/// </summary>
public abstract class RoomOperation
{
    public string Description { get; set; } = "";
    public DateTime Timestamp { get; set; } = DateTime.Now;

    /// <summary>
    /// 执行操作
    /// </summary>
    public abstract void Execute();

    /// <summary>
    /// 撤销操作
    /// </summary>
    public abstract void Undo();
}

/// <summary>
/// 修改房间名称操作
/// </summary>
public class RenameRoomOperation : RoomOperation
{
    private readonly RoomData _room;
    private readonly string _oldName;
    private readonly string _newName;

    public RenameRoomOperation(RoomData room, string newName)
    {
        _room = room;
        _oldName = room.Name;
        _newName = newName;
        Description = $"重命名: {_oldName} → {_newName}";
    }

    public override void Execute()
    {
        _room.Name = _newName;
    }

    public override void Undo()
    {
        _room.Name = _oldName;
    }
}

/// <summary>
/// 修改房间编号操作
/// </summary>
public class RenumberRoomOperation : RoomOperation
{
    private readonly RoomData _room;
    private readonly string _oldNumber;
    private readonly string _newNumber;

    public RenumberRoomOperation(RoomData room, string newNumber)
    {
        _room = room;
        _oldNumber = room.Number;
        _newNumber = newNumber;
        Description = $"重新编号: {_oldNumber} → {_newNumber}";
    }

    public override void Execute()
    {
        _room.Number = _newNumber;
    }

    public override void Undo()
    {
        _room.Number = _oldNumber;
    }
}

/// <summary>
/// 批量修改操作
/// </summary>
public class BatchModifyOperation : RoomOperation
{
    private readonly List<RoomOperation> _operations;

    public BatchModifyOperation(List<RoomOperation> operations)
    {
        _operations = operations;
        Description = $"批量修改 ({operations.Count} 项)";
    }

    public override void Execute()
    {
        foreach (var op in _operations)
        {
            op.Execute();
        }
    }

    public override void Undo()
    {
        // 逆序撤销
        for (int i = _operations.Count - 1; i >= 0; i--)
        {
            _operations[i].Undo();
        }
    }
}

/// <summary>
/// 修改参数操作
/// </summary>
public class SetParameterOperation : RoomOperation
{
    private readonly RoomData _room;
    private readonly string _parameterName;
    private readonly object? _oldValue;
    private readonly object? _newValue;

    public SetParameterOperation(RoomData room, string parameterName, object? newValue)
    {
        _room = room;
        _parameterName = parameterName;
        _oldValue = room.CustomParameters.TryGetValue(parameterName, out var val) ? val : null;
        _newValue = newValue;
        Description = $"设置参数: {parameterName} = {newValue}";
    }

    public override void Execute()
    {
        _room.CustomParameters[_parameterName] = _newValue;
    }

    public override void Undo()
    {
        if (_oldValue == null)
        {
            _room.CustomParameters.Remove(_parameterName);
        }
        else
        {
            _room.CustomParameters[_parameterName] = _oldValue;
        }
    }
}

/// <summary>
/// 操作事件参数
/// </summary>
public class OperationEventArgs : EventArgs
{
    public RoomOperation Operation { get; set; } = null!;
    public OperationType Type { get; set; }
}

/// <summary>
/// 操作类型
/// </summary>
public enum OperationType
{
    Execute,
    Undo,
    Redo
}
